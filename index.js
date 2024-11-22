import express from "express";
import mysql from "mysql2";
import PptxGenJS from "pptxgenjs";
import path from "path";
import { dirname } from "path";
import { fileURLToPath } from "url";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = process.env.PORT || 4000;

const connection = mysql.createConnection({
  host: process.env.HOST_DB,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
});
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

async function getData() {
  return new Promise((resolve, reject) => {
    connection.query("SELECT * FROM test", (err, results) => {
      if (err) {
        console.error("Error data:", err);
        reject(err);
      }
      resolve(results);
    });
  });
}

app.use(express.static(path.join(__dirname, "public")));

app.get("/generate", async (req, res) => {
  try {
    const slides = await getData();

    const pptx = new PptxGenJS();
    let yPosition = 1.5;
    const lineSpacing = 0.3;
    const {
      features,
      cuttingAngles,
      branding,
      image,
      series,
      models,
      productName,
    } = slides[0];

    //define slide
    pptx.defineSlideMaster({
      title: "TEMPLATE_SLIDE",
      margin: [0.5, 0.25, 1.0, 0.25],
      background: {
        color: "FFFFFF",
      },
      objects: [
        {
          rect: {
            x: 0,
            y: 0,
            w: "100%",
            h: 1.2,
            fill: { color: "DB011C" },
          },
        },
        {
          rect: {
            x: 1.6,
            y: "92%",
            w: 4.8,
            h: 0.01,
            fill: { color: "000000" },
          },
        },
        {
          text: {
            text: "Confidential document, property of TTI Group. For internal use only.",
            options: {
              x: 0.2,
              y: "94%",
              w: 5.5,
              h: 0.25,
              fontSize: 6,
            },
          },
        },
        {
          image: {
            path: "https://stg.milwaukeetool.asia/media/wysiwyg/page/job-apply/title-logo.png",
            x: 0.25,
            y: 0.3,
            w: 1.2,
            h: 0.7,
          },
        },
      ],
    });

    const slide = pptx.addSlide({ masterName: "TEMPLATE_SLIDE" });

    let startY = 1.4;
    let gap = 1.2;
    JSON.parse(cuttingAngles).forEach((item, index) => {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.3,
        y: startY + index * gap,
        w: 1.3,
        h: 1,
        line: { color: "000000", width: 1 },
        fill: { color: "FFFFFF" },
      });
      slide.addImage({
        path: item.img,
        x: 0.3,
        y: startY + index * gap,
        w: 1.3,
        h: 1,
      });
      slide.addText(`${item.angle} ${item.description}`, {
        x: 0.3,
        y: startY + index * gap + 1.1,
        w: 1.3,
        fontSize: 6,
        bold: true,
        color: "000000",
        align: "center",
      });
    });

    JSON.parse(features).forEach((feature) => {
      let lines = [];
      let text = feature;
      while (text.length > 0) {
        let line = text.substring(0, Math.min(4.8, text.length));
        lines.push(line);
        text = text.substring(line.length);
      }

      slide.addText(feature, {
        x: 1.6,
        y: yPosition,
        w: 4.8,
        fontSize: 7,
        color: "000000",
        fontFace: "Arial",
      });

      yPosition += lineSpacing;
    });

    slide.addText(productName, {
      y: 0.4,
      w: "100%",
      fontSize: 27,
      color: "ffffff",
      align: "right",
      fontFace: "Arial",
    });

    slide.addText(series, {
      y: 0.8,
      w: "100%",
      fontSize: 17,
      color: "ffffff",
      align: "right",
      fontFace: "Arial",
    });

    slide.addImage({
      path: JSON.parse(image).url,
      x: 6.4,
      y: 1.6,
      w: 3,
      h: 3,
      altText: JSON.parse(image).logo,
    });

    slide.addImage({
      path: JSON.parse(image).tag,
      x: 6.4,
      y: 1.6,
      w: 1.2,
      h: 0.2,
    });

    slide.addImage({
      path: JSON.parse(branding).logo,
      x: 1.3,
      y: 0.8,
      w: 0.4,
      h: 0.2,
      altText: JSON.parse(image).brand,
    });

    slide.addTable(JSON.parse(models), {
      x: 1.6,
      y: yPosition,
      w: 4.8,
      h: 1,
      colWidth: [2.5, 2, 2],
      fontSize: 6,
      align: "left",
      valign: "top",
    });

    const filePath = path.join(__dirname, "test.pptx");
    await pptx.writeFile({ fileName: filePath });
    res.download(filePath, "test.pptx");
  } catch (error) {
    console.error("Error:", error);
    res.status(500).send("Error create PowerPoint");
  }
});

app.listen(PORT, () => {
  console.log(`Server running http://localhost:${PORT}`);
});
