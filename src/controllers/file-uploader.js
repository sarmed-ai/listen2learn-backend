import multer from "multer";
import path from "path";
import fs from "fs";
import JSZip from "jszip";
import xml2js from "xml2js";
import sharp from "sharp";
import OpenAI from "openai";
import pLimit from "p-limit";

// Initialize OpenAI API with your API key
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

// Set up Multer storage for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads");
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  },
});

const upload = multer({ storage: storage }).single("file");

// Function to extract content and images from PPTX file
async function extractPptxContent(pptxPath) {
  const pptxStream = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(pptxStream);

  const slideRels = new Map();

  // Get slide relationship files
  const slideRelFiles = Object.keys(zip.files).filter((fileName) =>
    /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(fileName)
  );

  // Process slide relationships in parallel
  const slideRelPromises = slideRelFiles.map(async (relFileName) => {
    const relFile = zip.file(relFileName);
    if (!relFile) {
      console.error(`Relationship file not found: ${relFileName}`);
      return null;
    }
    const relXml = await relFile.async("string");
    const relData = await xml2js.parseStringPromise(relXml);
    const relationships = relData.Relationships.Relationship;
    const slideNumber = relFileName.match(/slide(\d+)\.xml\.rels$/)[1];
    slideRels.set(slideNumber, relationships);
  });

  await Promise.all(slideRelPromises);

  // Get slide files
  const slideFileNames = Object.keys(zip.files).filter((fileName) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(fileName)
  );

  const limit = pLimit(10); // Adjust the concurrency limit as needed

  const slidePromises = slideFileNames.map((slideFileName) =>
    limit(async () => {
      const slideFile = zip.file(slideFileName);
      if (!slideFile) {
        console.error(`Slide file not found: ${slideFileName}`);
        return null; // Skip this slide
      }
      const slideXml = await slideFile.async("string");
      const slideData = await xml2js.parseStringPromise(slideXml);
      const slideNumber = slideFileName.match(/slide(\d+)\.xml$/)[1];

      const slideContent = [];

      const shapes = slideData["p:sld"]["p:cSld"][0]["p:spTree"][0];

      // Extract text and images in order
      const elements = [].concat(shapes["p:sp"] || [], shapes["p:pic"] || []);

      for (const element of elements) {
        if (element["p:txBody"]) {
          // It's a text shape
          const paras = element["p:txBody"][0]["a:p"];
          let textContent = "";
          let i = 1;
          for (const para of paras) {
            if (para["a:r"]) {
              for (const run of para["a:r"]) {
                const texts = run["a:t"];
                if (texts) {
                  for (const text of texts) {
                    textContent += paras.length === 1 ? text : ` ${i}.` + text;
                  }
                }
              }
            }
            i += 1;
          }
          if (textContent.trim()) {
            slideContent.push({
              type: "text",
              content: textContent.trim(),
            });
          }
        } else if (element["p:blipFill"]) {
          // It's an image
          const blip = element["p:blipFill"][0]["a:blip"][0];
          const embedId = blip["$"]["r:embed"];
          const relationships = slideRels.get(slideNumber);
          if (!relationships) continue;
          const targetRel = relationships.find(
            (rel) => rel["$"].Id === embedId
          );
          if (!targetRel) continue;
          const target = targetRel["$"].Target;

          // Construct the image path
          let imagePath = target;
          if (!path.isAbsolute(imagePath)) {
            imagePath = path.posix.join(
              path.posix.dirname(`ppt/slides/slide${slideNumber}.xml`),
              target
            );
          }

          const imageFile = zip.file(imagePath);
          if (!imageFile) {
            console.error(`Image file not found: ${imagePath}`);
            continue;
          }

          // Extract the image file
          const imageData = await imageFile.async("nodebuffer");
          const imageFileName = `slide${slideNumber}_${path.basename(target)}`;
          const outputDir = path.join("output", "images");
          const outputPath = path.join(outputDir, imageFileName);

          // Ensure output directory exists
          fs.mkdirSync(outputDir, { recursive: true });

          const imageExt = path.extname(target).toLowerCase();
          try {
            if (imageExt === ".jpg" || imageExt === ".jpeg") {
              // If it's already a JPEG, no need to recompress
              await fs.promises.writeFile(outputPath, imageData);
            } else {
              // Convert to JPEG
              await sharp(imageData)
                .jpeg({ quality: 60 }) // Adjust quality as needed
                .toFile(outputPath);
            }
            slideContent.push({
              type: "image",
              content: outputPath,
            });
          } catch (err) {
            console.error("Error processing the image:", err);
            continue;
          }
        }
      }

      return { slideNumber: parseInt(slideNumber), slideContent };
    })
  );

  const slides = await Promise.all(slidePromises);

  // Remove any null slides that were skipped
  const validSlides = slides.filter((slide) => slide !== null);

  // Sort slides by slideNumber
  validSlides.sort((a, b) => a.slideNumber - b.slideNumber);

  return validSlides;
}

// Function to process slides and send content to OpenAI
const supportedFormats = new Set([".jpeg", ".jpg", ".png", ".tiff"]);

async function processSlides(slides, thread) {
  const messages = []; // Collect messages from all slides
  for (const slide of slides) {
    console.log(`\nProcessing Slide ${slide.slideNumber}`);

    // Process slide content elements in parallel
    const messageResultPromises = slide.slideContent.map(async (element) => {
      if (element.type === "text") {
        return {
          type: "text",
          text: element.content,
        };
      } else if (element.type === "image") {
        const ext = path.extname(element.content).toLowerCase();

        if (!supportedFormats.has(ext)) {
          console.log(`Unsupported image format: ${ext}`);
          return null;
        }

        // Process the image to get a text description
        const file = await getImageDescription(element.content);
        if (!file) {
          console.log(`Error processing image: ${element.content}`);
          return null;
        }

        return {
          type: "image_file",
          image_file: { file_id: file.id },
        };
      } else {
        return null;
      }
    });

    // Wait for all element processing to complete
    const messageResult = (await Promise.all(messageResultPromises)).filter(
      Boolean
    );
    // Create the message for the current slide
    if (
      messageResult.some((item) => item.text === "Learning Outcomes") ||
      slide.slideNumber === 1
    ) {
      continue;
    }

    const message = await openai.beta.threads.messages.create(thread, {
      role: "user",
      content: messageResult,
    });
    messages.push(message);
  }

  return messages;
}

// Function to get image description using Tesseract.js for OCR
async function getImageDescription(imagePath) {
  try {
    const supportedFormats = [".jpeg", ".jpg", ".png", ".tiff"];
    const ext = path.extname(imagePath).toLowerCase();

    if (!supportedFormats.includes(ext)) {
      console.log(`Unsupported image format: ${ext}`);
      return "Image Description: [Unsupported image format]";
    }

    // Log the image path to verify correctness
    console.log(`Processing image at: ${imagePath}`);

    const file = await openai.files.create({
      file: fs.createReadStream(imagePath),
      purpose: "vision",
    });
    return file;
  } catch (error) {
    console.error("Error processing image with Tesseract:", error);
    return "Image Description: [Error processing image]";
  }
}

const deleteFiles = async (messages) => {
  // Flatten the array and extract only the files of type 'image_file'
  const imageFiles = messages.flatMap((file) =>
    file.content.filter((innerFile) => innerFile.type === "image_file")
  );

  // Create an array of deletion promises
  const deletePromises = imageFiles.map((innerFile) =>
    openai.files.del(innerFile.image_file.file_id)
  );

  // Wait for all deletion promises to resolve in parallel
  await Promise.all(deletePromises);
};
// Exported function to handle file upload
export const fileUpload = async (req, res, next) => {
  try {
    upload(req, res, async function (err) {
      if (err) {
        return res.status(500).json({ error: err.message });
      }
      if (!req.file) {
        return res.status(400).json({ message: "No file uploaded" });
      }

      const pptFile = path.resolve(req.file.path);
      const assistant = "asst_RWO3Vnbk7CIGBx7A7ppmJeWa";
      const thread = await openai.beta.threads.create();

      try {
        const slides = await extractPptxContent(pptFile);
        const messages = await processSlides(slides, thread.id);

        let run = await openai.beta.threads.runs.createAndPoll(thread.id, {
          assistant_id: assistant,
        });
        const result = await openai.beta.threads.messages.list(run.thread_id);
        await deleteFiles(messages);
        const json = result.data.find((item) => item.role === "assistant")
          .content[0].text?.value;
        const output =
          typeof JSON.parse(json) === "object"
            ? JSON.parse(json)
            : { error: "Something went wrong!" };
        res.status(200).json({
          message: output,
          messages,
        });
      } catch (loaderError) {
        console.error(loaderError);
        res.status(400).json({ message: loaderError.message });
      }
    });
  } catch (error) {
    console.error(error);
    next(error);
  }
};
