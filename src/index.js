import express from "express";
import { fileUploader } from "./routes/index.js"; 
import 'dotenv/config'
import cors from "cors";
import path from "path";
import { fileURLToPath } from 'url';

const app = express();
const port = 3000;


const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

var corsOptions = {
  origin: '*',
  optionsSuccessStatus: 200 // some legacy browsers (IE11, various SmartTVs) choke on 204
}

app.use(express.json());
app.use(cors(corsOptions));
app.use(express.urlencoded({ extended: true }));
app.use("/uploads", express.static(path.join(__dirname, "uploads")));
app.get("/", (req, res) => {
  res.send("Hello World!");
});
app.post("/", fileUploader);

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
