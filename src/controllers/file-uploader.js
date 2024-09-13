import { PPTXLoader } from "@langchain/community/document_loaders/fs/pptx";
import { OpenAI } from "@langchain/openai";
import { readPptx } from "../utils/helpers.js";
import multer from "multer";
import path from "path";

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    return cb(null, "uploads");
  },
  filename: (req, file, cb) => {
    return cb(null, `${Date.now()}-${file.originalname}`);
  },
});

const upload = multer({ storage: storage }).single("file");

export const fileUpload = async (req, res, next) => {
  try {
    upload(req, res, async function (err) {
      if (err) {
        return res.status(500).json({ error: err.message });
      }
      if (!req.file) {
        return res.status(400).json({ message: "No file uploaded" });
      }

      const llm = new OpenAI({
        model: "gpt-4o",
        temperature: 0,
        maxTokens: undefined,
        timeout: undefined,
        maxRetries: 2,
        apiKey: `${process.env.OPENAI_API_KEY}`,
      });
      const pptFile = path.resolve(req.file.path);

      try {
        const loader = new PPTXLoader(pptFile);
        const docs = await loader.load();
        const inputText = `
        Guideline: Deliver a detailed and accurate explanation of the topic, covering all key points without skipping any important information. Use technical language and complex terms as needed, assuming the audience has a basic understanding. Ensure the explanation is factual and based on the provided context, without introducing any additional or speculative information. The response should be in plain text only, without any bullet points, titles, or headings. Since this will be directly converted into speech, be careful not to include anything inappropriate or offensive. Focus on clarity and precision, without simplifications or conclusions
        context: ${docs[0].pageContent}
        
        `;
        const completion = await llm.invoke(inputText);
        console.log(completion, "complete");
        res
          .status(200)
          .json({ message: "File uploaded successfully", file: req.file });
      } catch (loaderError) {
        console.error(loaderError);
        res.status(400).json({ message: loaderError.message });
      }
    });
  } catch (error) {
    console.error(error);
    next(error); // Make sure 'next' is available in the middleware chain
  }
};
//   const llm = new OpenAI({
//     model: "gpt-3.5-turbo-instruct",
//     temperature: 0,
//     maxTokens: undefined,
//     timeout: undefined,
//     maxRetries: 2,
//     apiKey: `${process.env.OPENAI_API_KEY}`,
//   });
//   const inputText = "OpenAI is an AI company that ";
//   const completion = await llm.invoke(inputText);
//   const loader = new PPTXLoader(assets[0].uri);
//   const docs = await loader.load();
//   for (const element of docs) {
//     console.log(element);
//   }
