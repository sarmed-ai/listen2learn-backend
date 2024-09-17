import { ChatOpenAI } from "@langchain/openai";
import multer from "multer";
import path from "path";
import fs from "fs";
import JSZip from "jszip";
import xml2js from "xml2js";
import { HfInference } from "@huggingface/inference"; // Import Hugging Face Inference
import Tesseract from "tesseract.js";
import sharp from "sharp";
import OpenAI from "openai";

// Initialize OpenAI API with your API key
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

// Initialize Hugging Face Inference API
const hf = new HfInference(process.env.HUGGINGFACE_API_TOKEN);

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
  const data = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(data);

  console.log("Files in zip archive:", Object.keys(zip.files));

  const slidePromises = [];

  // Get slide relationships to map images
  const slideRels = {};
  const slideRelFiles = Object.keys(zip.files).filter((fileName) =>
    /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(fileName)
  );

  console.log("Slide relationship files:", slideRelFiles);

  for (const relFileName of slideRelFiles) {
    const relFile = zip.file(relFileName);
    if (!relFile) {
      console.error(`Relationship file not found: ${relFileName}`);
      continue;
    }
    const relXml = await relFile.async("string");
    const relData = await xml2js.parseStringPromise(relXml);
    const relationships = relData.Relationships.Relationship;
    const slideNumber = relFileName.match(/slide(\d+)\.xml\.rels$/)[1];
    slideRels[slideNumber] = relationships;
  }

  // Get slide files
  const slideFileNames = Object.keys(zip.files).filter((fileName) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(fileName)
  );

  console.log("Slide files:", slideFileNames);

  for (const slideFileName of slideFileNames) {
    slidePromises.push(
      (async () => {
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
            for (const para of paras) {
              if (para["a:r"]) {
                for (const run of para["a:r"]) {
                  const texts = run["a:t"];
                  if (texts) {
                    for (const text of texts) {
                      textContent += text;
                    }
                  }
                }
              }
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
            const relationships = slideRels[slideNumber];
            if (!relationships) continue;
            const targetRel = relationships.find(
              (rel) => rel["$"].Id === embedId
            );
            if (!targetRel) continue;
            const target = targetRel["$"].Target;

            // Construct the image path
            const imagePath = path.posix.join(
              "ppt/media",
              path.basename(target)
            );

            console.log(`Embed ID: ${embedId}`);
            console.log(`Target: ${target}`);
            console.log(`Image Path: ${imagePath}`);

            const imageFile = zip.file(imagePath);
            if (!imageFile) {
              console.error(`Image file not found: ${imagePath}`);
              continue;
            }

            // Extract the image file
            const imageData = await imageFile.async("nodebuffer");
            const imageFileName = `slide${slideNumber}_${path.basename(
              target
            )}`;
            const outputDir = path.join("output", "images");
            const outputPath = path.join(outputDir, imageFileName);

            // Ensure output directory exists
            fs.mkdirSync(outputDir, { recursive: true });

            sharp(imageData)
              .jpeg({ quality: 50 }) // You can adjust the quality here (1-100)
              .toFile(outputPath, (err, info) => {
                if (err) {
                  console.error("Error compressing the image:", err);
                  return;
                }
                slideContent.push({
                  type: "image",
                  content: outputPath,
                });
              });
          }
        }

        return { slideNumber: parseInt(slideNumber), slideContent };
      })()
    );
  }

  const slides = await Promise.all(slidePromises);

  // Remove any null slides that were skipped
  const validSlides = slides.filter((slide) => slide !== null);

  // Sort slides by slideNumber
  validSlides.sort((a, b) => a.slideNumber - b.slideNumber);

  return validSlides;
}

// Function to process slides and send content to OpenAI
async function processSlides(slides, thread) {
  const messages = []; // Collect messages from all slides

  for (const slide of slides) {
    console.log(`\nProcessing Slide ${slide.slideNumber}`);

    for (const element of slide.slideContent) {
      if (element.type === "text") {
        const message = await openai.beta.threads.messages.create(thread, {
          role: "user",
          content: element.content,
        });
        messages.push(message);
      } else if (element.type === "image") {
        const supportedFormats = [".jpeg", ".jpg", ".png", ".tiff"];
        const ext = path.extname(element.content).toLowerCase();

        if (!supportedFormats.includes(ext)) {
          console.log(`Unsupported image format: ${ext}`);
          continue;
        }
        // Process the image to get a text description
        const file = await getImageDescription(element.content);
        if (!file) {
          console.log(`Error processing image: ${element.content}`);
          continue;
        }
        const message = await openai.beta.threads.messages.create(thread, {
          role: "user",
          content: [
            {
              type: "image_file",
              image_file: { file_id: file.id },
            },
          ],
        });
        messages.push(message);
      }
    }
  }

  // Process the combined content with OpenAI
  // await processContentWithOpenAI(messages);
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

    // Read image into a buffer
    const imageBuffer = fs.readFileSync(imagePath);
    const base64Image = imageBuffer.toString("base64");
    const resizedImageBuffer = await sharp(imageBuffer)
      .resize(224, 224) // Resize image to 224x224 as expected by the model
      .toBuffer();
    const reducedBase64Image = resizedImageBuffer.toString("base64");
    // Chat
    // const llm = new ChatOpenAI({
    //   model: "gpt-4o",
    //   temperature: 0,
    // });
    // const aiMsg = await llm.invoke([
    //   {
    //     role: "system",
    //     content:
    //       "You are a helpful assistant that translates images to English. Describe the image content.",
    //   },
    //   {
    //     role: "user",
    //     content: "What is in the image?",
    //   },
    //   {
    //     role: "user",
    //     content: `[Image]`,
    //     attachments: [
    //       {
    //         type: "image",
    //         data: imageBuffer,
    //         format: ext.replace(".", ""), // Sending the image format like jpeg, png, etc.
    //       },
    //     ],
    //   },
    // ]);
    // console.log(aiMsg);
    // if (aiMsg && aiMsg.content) {
    //   return `Image Description: ${aiMsg.content}`;
    // } else {
    //   return "Image Description: [No description generated]";
    // }
    // Use Tesseract.js to process the image
    // const {
    //   data: { text },
    // } = await Tesseract.recognize(imageBuffer, "eng", {
    //   logger: (m) => console.log(m),
    // });

    // if (text) {
    //   return `Image Description: ${text}`;
    // } else {
    //   return "Image Description: [No description generated]";
    // }
    // Use Hugging Face Inference API for image captioning
    // const result = await hf.imageToText({
    //   data: imageBuffer,
    //   model: "Salesforce/blip2-opt-2.7b", // You can change the model if desired
    // });

    // const response = await fetch(
    //   "https://api-inference.huggingface.co/models/nlpconnect/vit-gpt2-image-captioning",
    //   {
    //     headers: {
    //       Authorization: `Bearer ${process.env.HUGGINGFACE_API_TOKEN}`,
    //       "Content-Type": "application/json",
    //     },
    //     method: "POST",
    //     body: resizedImageBuffer,
    //   }
    // );
    // const result = await response.json();
    // console.log(result);
    // if (result.generated_text) {
    //   return `Image Description: ${result.generated_text}`;
    // } else {
    //   return "Image Description: [No description generated]";
    // }
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

// Function to send content to OpenAI
async function processContentWithOpenAI(messages) {
  try {
    const response = await openai.chat.completions.create({
      model: "gpt-4",
      messages: messages,
    });

    console.log("OpenAI Response:", response.choices[0].message.content);
  } catch (error) {
    console.error(
      "OpenAI API error:",
      error.response ? error.response.data : error.message
    );
  }
}

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
        // console.log(messages);
        // const result = await openai.beta.threads.messages.list(thread.id);
        // if (run.status === "completed") {
        const result = await openai.beta.threads.messages.list(run.thread_id);
        for (const file of messages) {
          if (file.content[0]?.image_file?.file_id) {
            console.log(file.content[0]?.image_file?.file_id, file.content[0]);
            await openai.files.del(file.content[0]?.image_file?.file_id);
          }
        }
        res.status(200).json({
          message: result.data.find((item) => item.role === "assistant")
            .content[0].text,
        });
        // } else {
        //   console.log(run.status);
        // }
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
