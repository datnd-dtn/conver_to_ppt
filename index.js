import express from "express";
import PptxGenJS from "pptxgenjs";
import admin from "firebase-admin";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

import serviceAccount from "./demo16-7-firebase-adminsdk-k551k-32e2dc8eae.json" assert { type: "json" };
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: "https://demo16-7-default-rtdb.firebaseio.com",
});

const app = express();
const PORT = 3001;

app.use(express.static(path.join(__dirname, "public")));

app.get("/generate", async (req, res) => {
  try {
    const db = admin.database();
    const ref = db.ref("/");
    const snapshot = await ref.once("value");

    if (snapshot.exists()) {
      const data = snapshot.val();
      const formattedData = Object.keys(data).map((key) => ({
        title: key,
        value: data,
      }));
      //   const {slide1, slide2, slide3, slide4} = data;
      const pptx = new PptxGenJS();
      let yPosition = 1.5;

      //define slide
      pptx.defineSlideMaster({
        title: "TEMPLATE_SLIDE",
        margin: [0.5, 0.25, 1.0, 0.25],
        background: {
          path: "https://img.freepik.com/free-photo/colorful-background-with-alcohol-ink_24972-1282.jpg?t=st=1732087672~exp=1732091272~hmac=a529588749921bb82031ff0eed88f2f1b5a1db7106d589d16b09fb7da4529ccb&w=996",
        },
        objects: [
          {
            text: {
              text: "DTN-solution",
              options: {
                x: 0,
                y: 6.9,
                w: "100%",
                align: "center",
                color: "000000",
                fontSize: 12,
              },
            },
          },
          {
            image: {
              path: "https://www.dtn-e.com/wp-content/uploads/2022/03/logo-180x80.png",
              x: 9,
              y: 0.25,
              w: 0.6,
              h: 0.3,
            },
          },
        ],
        slideNumber: {
          x: 8.0,
          y: 7.0,
          fontSize: 10,
          color: "000000",
          align: "center",
        },
      });

      //Slide 1 Text
      const slide1 = pptx.addSlide({ masterName: "TEMPLATE_SLIDE" });
      const slide1Text = Object.entries(data.slide1).map(
        ([key, value]) => `${value}`
      );
      slide1.addText(slide1Text.join("\n"), {
        x: 0.5,
        y: 0.5,
        fontSize: 18,
        color: "000000",
      });

      //Slide 2 Image
      const slide2 = pptx.addSlide({ masterName: "TEMPLATE_SLIDE" });
      slide2.addText(data.slide2.title, {
        x: 0.5,
        y: 0.5,
        fontSize: 18,
        color: "000000",
      });
      slide2.addImage({ path: data.slide2.imagePath, x: 1, y: 1, w: 6, h: 3 });

      //Slide 3 Table
      const slide3 = pptx.addSlide({ masterName: "TEMPLATE_SLIDE" });
      const headerTable = Object.keys(data.slide3.tableData[0]);
      const bodyTable = data.slide3.tableData.map((row) => Object.values(row));
      const tableData = [headerTable, ...bodyTable];
      console.log(tableData);
      slide3.addText(data.slide3.title, {
        x: 0.5,
        y: 0.5,
        fontSize: 18,
        color: "000000",
      });
      slide3.addTable(tableData, {
        x: 1,
        y: 1,
        w: 6,
        border: { pt: 1, color: "000000" },
        fill: "F4B183",
        color: "000000",
        fontSize: 14,
      });

      //Slide 4 Chart
      const slide4 = pptx.addSlide({ masterName: "TEMPLATE_SLIDE" });
      slide4.addText(data.slide4.title, {
        x: 0.5,
        y: 0.5,
        fontSize: 18,
        color: "000000",
      });
      slide4.addChart(pptx.ChartType.line, data.slide4.chartData, {
        x: 1,
        y: 1,
        w: 8,
        h: 4,
      });

      const filePath = path.join(__dirname, "test.pptx");
      await pptx.writeFile({ fileName: filePath });

      res.download(filePath, "test.pptx");
    } else {
      res.status(404).send("No db");
    }
  } catch (error) {
    console.error("Error:", error);
    res.status(500).send("Lỗi khi tạo PowerPoint");
  }
});

app.listen(PORT, () => {
  console.log(`Server running http://localhost:${PORT}`);
});
