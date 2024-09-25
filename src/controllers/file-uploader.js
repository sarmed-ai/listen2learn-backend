import WebSocket, { WebSocketServer } from 'ws';
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

const wss = new WebSocketServer({ port: 8080 });

wss.on('connection', (ws) => {
  console.log('Client connected');
  ws.on('close', () => console.log('Client disconnected'));
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
    const relData = await xml2js.parseStringPromise(relXml, {
      explicitArray: false,
    });
    const relationships = relData.Relationships.Relationship;
    const slideNumber = relFileName.match(/slide(\d+)\.xml\.rels$/)[1];
    slideRels.set(
      slideNumber,
      Array.isArray(relationships) ? relationships : [relationships]
    );
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
      const slideData = await xml2js.parseStringPromise(slideXml, {
        explicitArray: false,
      });
      const slideNumber = slideFileName.match(/slide(\d+)\.xml$/)[1];

      const slideContent = [];

      const shapes = slideData["p:sld"]["p:cSld"]["p:spTree"];

      // Extract content recursively to handle nested elements
      extractContentFromShapes(
        shapes,
        slideContent,
        slideNumber,
        slideRels,
        zip
      );

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

function extractContentFromShapes(
  shapes,
  slideContent,
  slideNumber,
  slideRels,
  zip
) {
  if (!shapes) return;

  const elements = [].concat(
    shapes["p:sp"] || [],
    shapes["p:pic"] || [],
    shapes["p:grpSp"] || [],
    shapes["p:graphicFrame"] || [],
    shapes["p:cxnSp"] || [] // Connection shapes can also contain nested elements
  );

  elements.forEach((element) => {
    if (element["p:txBody"]) {
      // It's a text shape
      const paras = element["p:txBody"]["a:p"];
      let textContent = "";
      if (Array.isArray(paras)) {
        paras.forEach((para) => {
          textContent += extractTextFromPara(para);
        });
      } else {
        // Handle single paragraph case
        textContent += extractTextFromPara(paras);
      }
      if (textContent.trim()) {
        slideContent.push({
          type: "text",
          content: textContent.trim(),
        });
      }
    } else if (element["p:blipFill"]) {
      // It's an image
      extractImageFromElement(
        element,
        slideContent,
        slideNumber,
        slideRels,
        zip
      );
    } else if (element["p:grpSp"]) {
      // It's a group shape, recursively extract content
      extractContentFromShapes(
        element["p:spTree"],
        slideContent,
        slideNumber,
        slideRels,
        zip
      );
    } else {
      // Check other nested elements like p:graphicFrame, p:cxnSp
      for (const key in element) {
        if (element.hasOwnProperty(key) && typeof element[key] === "object") {
          extractContentFromShapes(
            element[key],
            slideContent,
            slideNumber,
            slideRels,
            zip
          );
        }
      }
    }
  });
}

function extractImageFromElement(
  element,
  slideContent,
  slideNumber,
  slideRels,
  zip
) {
  const blip = element["p:blipFill"]["a:blip"];
  const embedId = blip["$"]["r:embed"];
  const relationships = slideRels.get(slideNumber);
  if (!relationships) return;
  const targetRel = relationships.find((rel) => rel["$"].Id === embedId);
  if (!targetRel) return;
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
    return;
  }

  // Extract the image file
  imageFile
    .async("nodebuffer")
    .then((data) => {
      const imageFileName = `slide${slideNumber}_${path.basename(target)}`;
      const outputDir = path.join("output", "images");
      const outputPath = path.join(outputDir, imageFileName);

      // Ensure output directory exists
      fs.mkdirSync(outputDir, { recursive: true });

      const imageExt = path.extname(target).toLowerCase();
      return processImage(data, outputPath, imageExt, slideContent);
    })
    .catch((err) => {
      console.error("Error extracting image:", err);
    });
}

async function processImage(imageData, outputPath, imageExt, slideContent) {
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
  }
}

function extractTextFromPara(para) {
  let textContent = "";
  const runs = para["a:r"];
  if (Array.isArray(runs)) {
    runs.forEach((run) => {
      if (run["a:t"]) {
        textContent += run["a:t"] + " ";
      }
    });
  } else if (runs && runs["a:t"]) {
    // Single run
    textContent += runs["a:t"] + " ";
  }
  return textContent;
}

// Function to process slides and send content to OpenAI
const supportedFormats = new Set([".jpeg", ".jpg", ".png", ".tiff"]);

async function processSlides(slides) {
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

    messages.push(messageResult);
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


// const groupSlides = (slides, chunkSize) => {
//   const chunks = [];
//   for (let i = 0; i < slides.length; i += chunkSize) {
//     chunks.push(slides.slice(i, i + chunkSize));
//   }
//   return chunks;
// };

function groupSlides(slides, groupSize) {
  let slideGroups = [];
  for (let i = 0; i < slides.length; i += groupSize) {
    const group = slides.slice(i, i + groupSize);
    if (group.length > 0) {
      slideGroups.push(group); // Only push non-empty groups
    }
  }
  return slideGroups;
}


export const fileUpload = async (req, res, next) => {
  try {
    upload(req, res, async function (err) {
      if (err) return res.status(500).json({ error: err.message });
      if (!req.file) return res.status(400).json({ message: 'No file uploaded' });

      const pptFile = path.resolve(req.file.path);
      const assistant = 'asst_RWO3Vnbk7CIGBx7A7ppmJeWa';

      try {
        const slides = await extractPptxContent(pptFile);
        const slideGroups = groupSlides(slides, 2);  

        for (let i = 0; i < slideGroups.length; i++) {
          const group = slideGroups[i];

          // Process the group
          const messages = await processSlides(group);
          const thread = await openai.beta.threads.create();

          // Send messages to OpenAI
          for (const message of messages) {
            await openai.beta.threads.messages.create(thread.id, {
              role: 'user',
              content: message,
            });
          }

          const run = await openai.beta.threads.runs.createAndPoll(thread.id, {
            assistant_id: assistant,
          });
          const result = await openai.beta.threads.messages.list(run.thread_id);

          // Parse the response
          const output = result.data.find((msg) => msg.role === 'assistant').content[0].text?.value;
          const parsedOutput = JSON.parse(output);

          console.log("Parseddd outputtt: ", parsedOutput.transcript_segments);
          // Send the result to the connected WebSocket client
          wss.clients.forEach((client) => {
            if (client.readyState === WebSocket.OPEN) {
              client.send(JSON.stringify({
                groupNumber: i + 1,
                result: parsedOutput.transcript_segments
              }));
            }else{
              console.log("Client not connected");
            }
          });
          console.log("Partial response Sent");
          // Optionally delete files and clean up after each group
          await deleteFiles(result.data);
        }

        
        res.status(200).json({ message: 'Processing started' });

        wss.clients.forEach((client) => {
          if (client.readyState === WebSocket.OPEN) {
            client.close();
          }
        });

      } catch (error) {
        console.error(error);
        res.status(400).json({ message: error.message });
      }
    });
  } catch (error) {
    next(error);
  }
};