import express from "express";
import { configDotenv } from "dotenv";
import path from "path";
import multer from "multer";
import fs from "fs";
import { fileURLToPath } from "url";
import * as AsposeSlides from "asposeslidescloud";

// Load environment variables
configDotenv();
const app = express();
const PORT = process.env.PORT || 3000;

// Get __dirname equivalent in ES module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Configure Aspose Slides API
const slidesApi = new AsposeSlides.SlidesApi(
  process.env.ASPOSE_CLIENT_ID,
  process.env.ASPOSE_CLIENT_SECRET
);

// Middleware
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

// Multer setup for image upload
const upload = multer({ dest: "uploads/" });

// Render home page
app.get("/", (req, res) => {
  res.render("index", { error: null });
});

// Function to copy theme from Test.pptx to target presentation
async function applyThemeFromTestPptx(targetName) {
  const themePath = "Test.pptx";

  await slidesApi.copyMasterSlide(
    targetName,         // target presentation name
    themePath,          // cloneFrom: source presentation name
    1,                  // cloneFromPosition: index of the master slide to copy
    null,               // cloneFromPassword
    null,               // cloneFromStorage
    true,               // applyToAll (optional - true if you want to apply to all existing slides)
    null,               // password (if target is password-protected)
    null,               // folder (if stored in a specific folder)
    null                // storage (optional storage name)
  );
}

// POST route to create presentation
app.post("/create", upload.single("slideImage"), async (req, res) => {
  try {
    const { presentationName, slideCount } = req.body;
    const imagePath = req.file?.path;

    if (!presentationName || !presentationName.endsWith(".pptx")) {
      throw new Error(
        "Presentation name must be provided and must end with .pptx"
      );
    }

    const slidesCount = parseInt(slideCount);
    if (isNaN(slidesCount) || slidesCount <= 0) {
      throw new Error("Slide count must be a positive number");
    }

    const folder = "";
    const storage = null;

    // Delete existing file if it exists
    const existsResponse = await slidesApi.objectExists(
      presentationName,
      folder,
      storage
    );
    if (existsResponse.body.exists) {
      await slidesApi.deleteFile(presentationName, folder, storage);
    }

    // Create a new presentation
    await slidesApi.createPresentation(presentationName, folder, storage);

    // Apply theme from Test.pptx
    await applyThemeFromTestPptx(presentationName);

    // Add additional slides
    for (let i = 1; i < slidesCount; i++) {
      await slidesApi.createSlide(
        presentationName,
        null,
        null,
        folder,
        storage
      );
    }

    // Handle image upload and transition effect
    if (imagePath) {
      const imageBase64 = fs.readFileSync(imagePath, { encoding: "base64" });

      for (let i = 1; i <= slidesCount; i++) {
        // Add transition to each slide
        const slide = new AsposeSlides.Slide();
        const transition = new AsposeSlides.SlideShowTransition();
        transition.type = "Circle";
        transition.speed = "Medium";
        slide.slideShowTransition = transition;
        await slidesApi.updateSlide(
          presentationName,
          i,
          slide,
          folder,
          storage
        );

        // Add logo image to top-right
        const picFrame = new AsposeSlides.PictureFrame();
        picFrame.x = 620;
        picFrame.y = 10;
        picFrame.width = 52.5;
        picFrame.height = 52.5;

        const fillFormat = new AsposeSlides.PictureFill();
        fillFormat.base64Data = imageBase64;
        fillFormat.pictureFillMode = "Stretch";
        picFrame.pictureFillFormat = fillFormat;

        await slidesApi.createShape(
          presentationName,
          i,
          picFrame,
          null,
          null,
          null,
          folder
        );
      }

      // Remove uploaded image
      fs.unlinkSync(imagePath);
    }

    // Add animation shape to slide 1
    const shape = new AsposeSlides.Shape();
    shape.x = 200;
    shape.y = 150;
    shape.width = 200;
    shape.height = 100;
    shape.shapeType = "Rectangle";
    shape.text = "Click to Animate";
    shape.fillFormat = {
      type: "Solid",
      solidFillColor: { color: "#3498db" },
    };
    shape.lineFormat = {
      style: "Single",
      width: 2,
      dashStyle: "Solid",
    };

    const shapeResponse = await slidesApi.createShape(
      presentationName,
      1,
      shape,
      null,
      null,
      null,
      folder
    );

    const shapeHref = shapeResponse.body.selfUri.href;
    const shapeIndex = parseInt(shapeHref.split("/").pop(), 10);

    const animationEffect = {
      type: "Fly",
      subtype: "Bottom",
      triggerType: "OnClick",
      shapeIndex: shapeIndex,
    };

    await slidesApi.setAnimation(
      presentationName,
      1,
      { mainSequence: [animationEffect] },
      null,
      null,
      folder
    );

    const downloadUrl = `https://api.aspose.cloud/v3.0/slides/${presentationName}/download`;

    res.json({
      success: true,
      downloadUrl,
    });
  } catch (error) {
    console.error("❌ Error:", error.message);
    res.status(500).json({
      success: false,
      message: error.message || "Unknown error occurred",
    });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
