// index.js
import express from "express";
import { configDotenv } from "dotenv";
import path from "path";
import multer from "multer";
import fs from "fs";
import { fileURLToPath } from "url";
import * as AsposeSlides from "asposeslidescloud";
import { Console } from "console";

configDotenv();
const app = express();
const PORT = process.env.PORT || 3000;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const slidesApi = new AsposeSlides.SlidesApi(
  process.env.ASPOSE_CLIENT_ID,
  process.env.ASPOSE_CLIENT_SECRET
);

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ dest: "uploads/" });

app.get("/", (req, res) => {
  res.render("index", { error: null });
});

app.post("/create", upload.single("slideImage"), async (req, res) => {
  try {
    const { presentationName } = req.body;
    const imagePath = req.file?.path;

    if (!presentationName || !presentationName.endsWith(".pptx")) {
      throw new Error("Presentation name must end with .pptx");
    }

    const folder = "";
    const storage = "Test Storage";

    // Delete existing file if it exists
    const existsResponse = await slidesApi.objectExists(
      presentationName,
      folder
    );
    if (existsResponse.body.exists) {
      await slidesApi.deleteFile(presentationName, folder);
    }

    // Create a blank presentation
    await slidesApi.createPresentation(presentationName, null, null, folder);

    // Set slide properties
    for (let i = 1; i <= 3; i++) {
      await slidesApi.setSlideProperties(
        presentationName,
        {
          FirstSlideNumber: i,
          Orientation: "Landscape",
          ScaleType: "DoNotScale",
          SizeType: "WideScreen",
          Width: 960,
          Height: 720,
        },
        folder
      );
    }

    // Add additional slides
    await slidesApi.createSlide(presentationName, null, null, folder);

    const inchToPt = (inch) => inch * 72;

    // Page 1 Logic Added

    // Page 1 Handling

    // Add uploaded image
    if (imagePath) {
      const imageBase64 = fs.readFileSync(imagePath, { encoding: "base64" });

      for (let i = 1; i <= 2; i++) {
        const picFrame = new AsposeSlides.PictureFrame();
        picFrame.x = inchToPt(0);
        picFrame.y = inchToPt(0);
        picFrame.width = inchToPt(4.27);
        picFrame.height = inchToPt(7.5);

        const fillFormat = new AsposeSlides.PictureFill();
        fillFormat.base64Data = imageBase64;
        fillFormat.pictureFillMode = "Stretch";
        picFrame.pictureFillFormat = fillFormat;

        await slidesApi.createShape(
          presentationName,
          1,
          picFrame,
          null,
          null,
          null,
          folder
        );
      }

      fs.unlinkSync(imagePath); // Delete local image after use
    }

    // Add rectangles
    const rectangles = [
      { x: 0, y: 1.89, width: 0.71, height: 2.14 },
      { x: 0, y: 6.68, width: 2.26, height: 0.82 },
      { x: 4.28, y: 0, width: 9.07, height: 4.4 },
    ];

    for (const r of rectangles) {
      await slidesApi.createShape(
        presentationName,
        1,
        {
          shapeType: "Rectangle",
          x: inchToPt(r.x),
          y: inchToPt(r.y),
          width: inchToPt(r.width),
          height: inchToPt(r.height),
          text: "",
          fillFormat: {
            type: "Solid",
            color: "#FFFFCA08",
          },
          lineFormat: {
            type: "Solid",
            style: "Single",
            width: 0,
            fillFormat: {
              type: "Solid",
              solidFillColor: {
                color: "#FFFFCA08", // ✅ Correct hex format
              },
            },
          },
        },
        null,
        null,
        null,
        folder
      );
    }

    const animationEffects = [];

    // Rectangle text box definitions
    const textBoxes = [
      {
        text: "Title Overview",
        x: 5.53,
        y: 0.33,
        width: 7.07,
        height: 3.95,
        fontSize: 54,
        bold: true,
        alignH: "Left",
      },
      {
        text: "Enter Overview Details in a Form of Heading",
        x: 5.53,
        y: 4.61,
        width: 7.07,
        height: 1.63,
        fontSize: 24,
        bold: false,
        alignH: "Left",
      },
      {
        text: "Presenter Name",
        x: 5.54,
        y: 6.68,
        width: 2.3,
        height: 0.33,
        fontSize: 16,
        bold: false,
        alignH: "Left",
      },
    ];

    for (const [index, box] of textBoxes.entries()) {
      // Step 1: Create basic shape with transparent outline
      const shapeResponse = await slidesApi.createShape(
        presentationName,
        1,
        {
          shapeType: "Rectangle",
          x: inchToPt(box.x),
          y: inchToPt(box.y),
          width: inchToPt(box.width),
          height: inchToPt(box.height),
          fillFormat: { type: "NoFill" },
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000", // Fully transparent ARGB
            },
            width: 0,
          },
          text: box.text,
          paragraphs: [
            {
              alignment: box.alignH,
            },
          ],
        },
        null,
        null,
        null,
        folder
      );

      const shapeData = shapeResponse.body;

      let shapeIndex;
      if (shapeData?.index !== undefined) {
        shapeIndex = shapeData.index;
      } else if (shapeData?.selfUri?.href) {
        const match = shapeData.selfUri.href.match(/shapes\/(\d+)/);
        if (match) {
          shapeIndex = parseInt(match[1], 10);
        }
      }

      if (shapeIndex === undefined) {
        throw new Error("❌ Unable to determine shape index from response!");
      }

      if (index === 1 || index === 2) {
        animationEffects.push({
          type: "Fly",
          subtype: index === 1 ? "Bottom" : "Right",
          triggerType: "onClick",
          shapeIndex: shapeIndex,
          presetClassType: "Entrance",
          acceleration: 0.1,
          duration: 1,
        });
      }

      // Step 2: Update text frame properties for vertical alignment
      await slidesApi.updateShape(
        presentationName,
        1,
        shapeIndex,
        {
          textFrameFormat: {
            anchoringType:
              index === 0 ? "Bottom" : index === 1 ? "Top" : "Center",
          },
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
            width: 0,
          },
        },
        folder
      );

      // Step 3: Update portion formatting
      await slidesApi.updatePortion(
        presentationName,
        1,
        shapeIndex,
        1,
        1,
        {
          text: box.text,
          fontHeight: box.fontSize,
          latinFont: "Arial",
          fontColor: "#FF000000",
          fontBold: box.bold ? "True" : "False",
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
            width: 0,
          },
          mathParagraph: {
            justification: "LeftJustified",
          },
        },
        folder
      );
    }
    if (animationEffects.length > 0) {
      await slidesApi.setAnimation(
        presentationName,
        1,
        { mainSequence: animationEffects },
        null,
        null,
        folder
      );
    }

    // Page 2 Handling

    const page2AnimationEffects = [];

    // RECTANGLES FOR PAGE 2 -- PART 1
    const rectanglesPage2 = [
      {
        x: 0,
        y: 0,
        width: 8.5,
        height: 0.38,
        color: "#FFFFCA08", // Top 1st Rectangle
      },
      {
        x: 8.4,
        y: 0,
        width: 4.95,
        height: 0.38,
        color: "#FFFFDF6B", // Top 2nd Rectangle
      },
      {
        x: 0,
        y: 7.12,
        width: 4.95,
        height: 0.38,
        color: "#FFFFDF6B", // Bottom 3rd Rectangle
      },
      {
        x: 4.95,
        y: 7.12,
        width: 8.4,
        height: 0.38,
        color: "#FFFFCA08", // Bottom 4th Rectangle
      },
    ];

    for (const rect of rectanglesPage2) {
      await slidesApi.createShape(
        presentationName,
        2, // <-- Page 2
        {
          shapeType: "Rectangle",
          x: inchToPt(rect.x),
          y: inchToPt(rect.y),
          width: inchToPt(rect.width),
          height: inchToPt(rect.height),
          text: "",
          fillFormat: {
            type: "Solid",
            color: rect.color,
          },
          lineFormat: {
            type: "Solid",
            style: "Single",
            width: 0,
            fillFormat: {
              type: "Solid",
              solidFillColor: {
                color: rect.color,
              },
            },
          },
        },
        null,
        null,
        null,
        folder
      );
    }

    // --- Part 2 of Page 2 ---
    // (A) Main Horizontal Rectangle on Page 2
    await slidesApi.createShape(
      presentationName,
      2,
      {
        shapeType: "Rectangle",
        x: inchToPt(0),
        y: inchToPt(2.08),
        width: inchToPt(13.34),
        height: inchToPt(0.12),
        fillFormat: {
          type: "Solid",
          color: "#FFF2F2F2",
        },
        lineFormat: {
          type: "Solid",
          width: 0,
          fillFormat: {
            type: "Solid",
            color: "#FFF2F2F2",
          },
        },
        text: "",
      },
      null,
      null,
      null,
      folder
    );

    // (B) 4 Circles (ShapeType: "Oval")
    const circles = [
      { x: 2.29, y: 2.05, color: "#FF9641E7" }, // A
      { x: 5.18, y: 2.05, color: "#FFEF476B" }, // B
      { x: 7.98, y: 2.05, color: "#FF9641E7" }, // C
      { x: 10.88, y: 2.05, color: "#FFEF476B" }, // D
    ];

    for (const c of circles) {
      await slidesApi.createShape(
        presentationName,
        2,
        {
          shapeType: "Ellipse",
          x: inchToPt(c.x),
          y: inchToPt(c.y),
          width: inchToPt(0.2),
          height: inchToPt(0.2),
          fillFormat: {
            type: "Solid",
            color: c.color,
          },
          lineFormat: {
            type: "Solid",
            width: 0,
            fillFormat: {
              type: "Solid",
              color: c.color,
            },
          },
          text: "",
        },
        null,
        null,
        null,
        folder
      );
    }

    // (C) 4 Diamonds (ShapeType: "Diamond")
    const diamonds = [
      { x: 1.95, y: 2.21, color: "#FF9641E7" }, // Rhombus A
      { x: 4.84, y: 2.21, color: "#FFEF476B" }, // Rhombus B
      { x: 7.64, y: 2.21, color: "#FF9641E7" }, // Rhombus C
      { x: 10.53, y: 2.21, color: "#FFEF476B" }, // Rhombus D
    ];

    for (const d of diamonds) {
      await slidesApi.createShape(
        presentationName,
        2,
        {
          shapeType: "Diamond",
          x: inchToPt(d.x),
          y: inchToPt(d.y),
          width: inchToPt(0.88),
          height: inchToPt(0.88),
          fillFormat: {
            type: "Solid",
            color: d.color,
          },
          lineFormat: {
            type: "Solid",
            width: 0,
            fillFormat: {
              type: "Solid",
              color: d.color,
            },
          },
          text: "",
        },
        null,
        null,
        null,
        folder
      );
    }

    const iconPositions = [
      { left: 2.12, top: 2.37, file: "icon/Icon1.ico" },
      { left: 4.94, top: 2.37, file: "icon/Icon2.ico" },
      { left: 7.8, top: 2.37, file: "icon/Icon3.ico" },
      { left: 10.69, top: 2.37, file: "icon/Icon4.ico" },
    ];

    const EMPTY_FILE_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

    for (const icon of iconPositions) {
      if (!fs.existsSync(icon.file)) {
        throw new Error(`Icon file not found: ${icon.file}`);
      }

      const iconBase64 = fs.readFileSync(icon.file, { encoding: "base64" });

      const oleFrame = new AsposeSlides.OleObjectFrame();

      // Required Position/Size
      oleFrame.x = inchToPt(icon.left);
      oleFrame.y = inchToPt(icon.top);
      oleFrame.width = inchToPt(0.53);
      oleFrame.height = inchToPt(0.53);

      // Mandatory OLE Properties
      oleFrame.embeddedFileBase64Data = EMPTY_FILE_BASE64;
      oleFrame.embeddedFileExtension = "png";
      oleFrame.objectProgId = "Paint.Picture"; // Changed to valid ProgID

      // Icon Configuration
      oleFrame.isObjectIcon = true;
      oleFrame.substitutePictureFormat = new AsposeSlides.PictureFill({
        base64Data: iconBase64,
        pictureFillMode: AsposeSlides.PictureFill.PictureFillModeEnum.Stretch,
      });
      oleFrame.substitutePictureTitle = "Icon Preview"; // Required title

      // Line Format Configuration
      oleFrame.lineFormat = new AsposeSlides.LineFormat({
        dashStyle: AsposeSlides.LineFormat.DashStyleEnum.Solid,
        width: 0,
        fillFormat: new AsposeSlides.SolidFill({
          color: "#00000000",
        }),
      });

      await slidesApi.createShape(
        presentationName,
        2,
        oleFrame,
        null,
        null,
        null,
        folder
      );
    }
    // Page 2 - Part 3
    const part3Rectangles = [
      { x: 1.05, y: 3.36 },
      { x: 3.93, y: 3.36 },
      { x: 6.76, y: 3.36 },
      { x: 9.64, y: 3.36 },
    ];

    for (const { x, y } of part3Rectangles) {
      const shapeRes = await slidesApi.createShape(
        presentationName,
        2,
        {
          shapeType: "Rectangle",
          x: inchToPt(x),
          y: inchToPt(y),
          width: inchToPt(2.66),
          height: inchToPt(2.79),
          fillFormat: {
            type: "Solid",
            color: "#FFF2F2F2",
          },
          lineFormat: {
            type: "Solid",
            width: 0,
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
          },
        },
        null,
        null,
        null,
        folder
      );
      const shapeData = shapeRes.body;
      let shapeIndex;
      if (shapeData?.index !== undefined) {
        shapeIndex = shapeData.index;
      } else if (shapeData?.selfUri?.href) {
        const match = shapeData.selfUri.href.match(/shapes\/(\d+)/);
        if (match) shapeIndex = parseInt(match[1], 10);
      }

      if (shapeIndex === undefined) {
        throw new Error("❌ Unable to determine shape index for title box.");
      }

      page2AnimationEffects.push(shapeIndex);
    }

    // Title TextBoxes
    const titleBoxes = [
      { x: 1.07, y: 3.55 },
      { x: 3.93, y: 3.55 },
      { x: 6.76, y: 3.55 },
      { x: 9.64, y: 3.55 },
    ];

    for (const { x, y } of titleBoxes) {
      const titleText = "Enter Title Here";

      // Step 1: Create shape with initial text and proper alignment
      const shapeRes = await slidesApi.createShape(
        presentationName,
        2,
        {
          shapeType: "Rectangle",
          x: inchToPt(x),
          y: inchToPt(y),
          width: inchToPt(2.66),
          height: inchToPt(0.44),
          fillFormat: { type: "NoFill" },
          lineFormat: {
            type: "Solid",
            width: 0,
            fillFormat: { type: "Solid", color: "#00000000" },
          },
          text: titleText,
          paragraphs: [{ alignment: "Center" }],
        },
        null,
        null,
        null,
        folder
      );

      const shapeData = shapeRes.body;

      let shapeIndex;
      if (shapeData?.index !== undefined) {
        shapeIndex = shapeData.index;
      } else if (shapeData?.selfUri?.href) {
        const match = shapeData.selfUri.href.match(/shapes\/(\d+)/);
        if (match) shapeIndex = parseInt(match[1], 10);
      }

      if (shapeIndex === undefined) {
        throw new Error("❌ Unable to determine shape index for title box.");
      }

      // Step 2: Update text frame (anchoring + line format)
      await slidesApi.updateShape(
        presentationName,
        2,
        shapeIndex,
        {
          textFrameFormat: {
            anchoringType: "Center",
          },
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
            width: 0,
          },
        },
        folder
      );

      // Step 3: Update portion formatting (most important)
      await slidesApi.updatePortion(
        presentationName,
        2,
        shapeIndex,
        1,
        1,
        {
          text: titleText,
          fontHeight: 20,
          latinFont: "Arial",
          fontBold: "True",
          fontColor: "#FF000000",
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
            width: 0,
          },
        },
        folder
      );
      page2AnimationEffects.push(shapeIndex);
    }

    // Paragraph TextBoxes
    const paraBoxes = [
      { x: 1.37, y: 4.11 },
      { x: 4.28, y: 4.11 },
      { x: 7.11, y: 4.11 },
      { x: 9.99, y: 4.11 },
    ];

    for (const { x, y } of paraBoxes) {
      const paraText = "Paragraph for the description is placed here.....";

      // Step 1: Create shape
      const shapeRes = await slidesApi.createShape(
        presentationName,
        2,
        {
          shapeType: "Rectangle",
          x: inchToPt(x),
          y: inchToPt(y),
          width: inchToPt(1.95),
          height: inchToPt(0.81),
          fillFormat: { type: "NoFill" },
          lineFormat: {
            type: "Solid",
            width: 0,
            fillFormat: { type: "Solid", color: "#00000000" },
          },
          text: paraText,
          paragraphs: [{ alignment: "Left" }],
        },
        null,
        null,
        null,
        folder
      );

      const shapeData = shapeRes.body;

      // Step 2: Get shape index
      let shapeIndex;
      if (shapeData?.index !== undefined) {
        shapeIndex = shapeData.index;
      } else if (shapeData?.selfUri?.href) {
        const match = shapeData.selfUri.href.match(/shapes\/(\d+)/);
        if (match) shapeIndex = parseInt(match[1], 10);
      }

      if (shapeIndex === undefined) {
        throw new Error(
          "❌ Unable to determine shape index for paragraph box."
        );
      }

      // Step 3: Update text frame alignment and transparency
      await slidesApi.updateShape(
        presentationName,
        2,
        shapeIndex,
        {
          textFrameFormat: {
            anchoringType: "Top",
          },
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
            width: 0,
          },
        },
        folder
      );

      // Step 4: Update text portion styling
      await slidesApi.updatePortion(
        presentationName,
        2,
        shapeIndex,
        1,
        1,
        {
          text: paraText,
          fontHeight: 14,
          latinFont: "Arial",
          fontColor: "#FF000000",
          fontBold: "False",
          lineFormat: {
            type: "Solid",
            fillFormat: {
              type: "Solid",
              color: "#00000000",
            },
            width: 0,
          },
          mathParagraph: {
            justification: "LeftJustified",
          },
        },
        folder
      );
      page2AnimationEffects.push(shapeIndex);
    }

    console.log("Page 2 Animation Effects:", page2AnimationEffects);

    if (page2AnimationEffects.length > 0) {
      const animationSequence = page2AnimationEffects.map(
        (shapeIndex, idx) => ({
          type: "Bounce",
          triggerType: idx === 0 ? "OnClick" : "WithPrevious",
          shapeIndex: shapeIndex,
          presetClassType: "Entrance",
          acceleration: 0.1,
          duration: 0.5,
        })
      );

      await slidesApi.setAnimation(
        presentationName,
        2, // Page 2
        { mainSequence: animationSequence },
        null,
        null,
        folder
      );
    }

    // Add Logo Logic Added
    // Add default logo if it exists
    const logoPath = path.join(__dirname, "public", "images", "Xebia_Logo.jpg");
    if (fs.existsSync(logoPath)) {
      const logoBase64 = fs.readFileSync(logoPath, { encoding: "base64" });

      const picFrame = new AsposeSlides.PictureFrame();
      picFrame.x = inchToPt(0.1);
      picFrame.y = inchToPt(0.1);
      picFrame.width = inchToPt(1.03);
      picFrame.height = inchToPt(0.9);

      const fillFormat = new AsposeSlides.PictureFill();
      fillFormat.base64Data = logoBase64;
      fillFormat.pictureFillMode = "Stretch";
      picFrame.pictureFillFormat = fillFormat;

      for (let i = 1; i <= 2; i++) {
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
    }

    const downloadUrl = `https://api.aspose.cloud/v3.0/slides/${presentationName}/download`;

    res.json({ success: true, downloadUrl });
  } catch (error) {
    console.error("❌ Error:", error.message);
    res.status(500).json({ success: false, message: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
