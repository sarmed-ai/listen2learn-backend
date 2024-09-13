async function readPptx(filePath) {
  const pptx = new PptxGenJS();
  const fileBuffer = fs.readFileSync(filePath);

  // Load the PPTX file
  await pptx.load(fileBuffer);

  // Access slides and their content
  pptx.getSlides().forEach((slide, index) => {
    console.log(`Slide ${index + 1}:`);
    slide.getShapes().forEach((shape) => {
      console.log(shape.text);
    });
  });
}

export { readPptx };
