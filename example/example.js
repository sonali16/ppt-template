
const TEMPLATE = 'TestInput1.pptx';
const OUTPUT = 'TestIntermediateOutput1.pptx';
const FINALOUTPUT = 'TestOutput1.pptx';
const imagePath = 'newImage.jpeg'; // Path to the image file to be added

var PPT_Template = require('../');
const fs = require('fs');
const JSZip = require('jszip');
const PPTX = require('nodejs-pptx');
var Presentation = PPT_Template.Presentation;
var Slide = PPT_Template.Slide;


// Function to get position and size of images in a PowerPoint slide
async function getImagePositionsAndSizes(presentationPath) {
    const pptxData = await fs.promises.readFile(presentationPath);
    const zip = await JSZip.loadAsync(pptxData);
    const slidePaths = Object.keys(zip.files).filter(fileName => fileName.startsWith('ppt/slides/slide'));
    console.log('slidePaths:', slidePaths);
    const imageInfo = [];

    for (const slidePath of slidePaths) {
        const slideData = await zip.file(slidePath).async('string');
        
        // Regular expression to match image tags in the slide XML content
        const imageRegex = /<p:pic>(.*?)<\/p:pic>/gs;
        let match;
        while ((match = imageRegex.exec(slideData)) !== null) {
            const imageData = match[1];
            console.log('match1:', imageData);
            
            // Extracting position and size attributes from the image tag
            const xMatch = /<a:xfrm>(?:.|\n)*?<a:off x="(.*?)".*?\/>/s.exec(imageData);
            console.log('xMatch:', xMatch);
            const yMatch = /<a:xfrm>(?:.|\n)*?<a:off(?:.|\n)*? y="(.*?)".*?\/>/s.exec(imageData);
            console.log('yMatch:', yMatch);
            const widthMatch = /<a:ext cx="(.*?)".*?>/s.exec(imageData);
            console.log('widthMatch:', widthMatch);
            const heightMatch = /<a:ext(?:.|\n)*? cy="(.*?)".*?>/s.exec(imageData);
            console.log('heightMatch:', heightMatch);
            
            if (xMatch && yMatch && widthMatch && heightMatch) {
                const x = parseFloat(xMatch[1]) / 12700; // converting EMU to points (1 inch = 72 points, 1 inch = 12700 EMU)
                const y = parseFloat(yMatch[1]) / 12700;
                const width = parseFloat(widthMatch[1]) / 12700;
                const height = parseFloat(heightMatch[1]) / 12700;

                imageInfo.push({
                    path: slidePath,
                    x,
                    y,
                    width,
                    height
                });
            }
        }
    }

    return imageInfo;
}

async function addImage(imagePath, x, y, cx, cy) {
    let pptx = new PPTX.Composer();

await pptx.load(OUTPUT);
await pptx.compose(async pres => {
   // console.log('pres:', pres);
  await pres.getSlide('slide1').addImage(image => {
    image
      .file(imagePath)
      .x(x)
      .y(y)
      .cx(cx)
      .cy(cy);
      console.log('image added:', image);
  });
});
await pptx.save(FINALOUTPUT);
}


// Presentation Object
var myPresentation = new Presentation();

console.log('# Load test.pptx as template, then build output.pptx with custom content.');

// Load example.pptx
myPresentation.loadFile(TEMPLATE)

	.then(() => {
		console.log('- Read Presentation File Successfully!');
	})

	.then(() => {

		// get slide conut
		var slideCount = myPresentation.getSlideCount();
		console.log('- Slides Count is ', slideCount);

		// Get slide by index. (Base from 1)
		var slideIndex1 = 1;

		// Get and clone slide. (Watch out index...)
		let cloneSlide1 = myPresentation.getSlide(slideIndex1).clone();

		// Fill all content
		cloneSlide1.fillAll([
			Slide.pair('[TITLE]', 'Hello PPT'),
		]);


		// Generate new presention by silde array.
		var newSlides = [cloneSlide1];
		return myPresentation.generate(newSlides);
	})

	.then((newPresentation) => {
		console.log('- Generate Presentation Successfully');
		return newPresentation;
	})

	.then((newPresentation) => {
		// Output .pptx file
		return newPresentation.saveAs(OUTPUT);
	})

	.then(() => {
		console.log('- Save Successfully');
		getImagePositionsAndSizes(OUTPUT)
    .then((imageInfo) => {
        console.log('Image positions and sizes:', imageInfo);
        addImage(imagePath, imageInfo[0].x, imageInfo[0].y, imageInfo[0].width, imageInfo[0].height);
    })
    .catch((error) => {
        console.error('Error:', error);
    });
	})

	.catch((err) => {
		console.error(err);
	});