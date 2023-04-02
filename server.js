const prompt = require("prompt-sync")();
const express = require("express");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const bodyParser = require("body-parser");

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res, next) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

var wardens = [
  "Krishna M",
  "Arun Kumar Jalan",
  "Nitin Chaturvedi",
  "Krishnendra Shekhawat",
  "Rakhee",
  "Kumar Sankar Bhattacharya",
  "Praveen Kumar A.V.",
  "Dipendu Bhunia",
  "Sharad Shrivastava",
];
var hostels = [
  "Srinivasa Ramanujan Bhawan",
  "Krishna Bhawan",
  "Gandhi Bhawan",
  "Vishwakarma Bhawan",
  "Meera Bhawan",
  "Shankar Bhawan",
  "Vyas Bhawan",
  "Ram Bhawan",
  "Budh Bhawan",
];

app.post("/leave", (req, res, next) => {
  var warden_name = wardens[hostels.indexOf(req.body.bhawan)];

  const content = fs.readFileSync(
    path.resolve(__dirname, "template.docx"),
    "binary"
  );

  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render({
    id: req.body.id,
    full_name: req.body.full_name,
    phone: req.body.phone,
    bhawan: req.body.bhawan,
    room: req.body.room,
    warden_name: warden_name,
    dep_date:
      req.body.dep_date + "-" + req.body.dep_month + "-" + req.body.dep_year,
    ret_date:
      req.body.ret_date + "-" + req.body.ret_month + "-" + req.body.ret_year,
  });

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  res.setHeader("Content-Disposition", "attachment; filename=mydocument.docx");
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
  res.setHeader("Content-Length", buf.length);

  res.send(buf);
});

app.listen(3000);
