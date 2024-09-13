import express from "express";
import { fileUpload } from "../controllers/file-uploader.js";
const router = express.Router();

router.post("/", fileUpload);

export default router;
