let singleCertificateForm = document.getElementById("single-certificate-form");
let multipleCertificateForm = document.getElementById("multiple-certificate-form");
let nameInput = document.getElementById("name");
let fileInput = document.getElementById("file");

const capitalize = (str, lower = false) =>
  (lower ? str.toLowerCase() : str).replace(/(?:^|\s|["'([{])+\S/g, (match) =>
    match.toUpperCase()
  );

singleCertificateForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const submitButton = e.target.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = "Generating...";

  if (nameInput.value.trim()) {
    await handleSingleName(nameInput.value.trim());
  } else {
    alert("Please enter a name.");
  }

  submitButton.disabled = false;
  submitButton.textContent = "Generate Single Certificate";
  singleCertificateForm.reset();
});

multipleCertificateForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const submitButton = e.target.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = "Generating...";

  if (fileInput.files.length > 0) {
    await handleExcelFile(fileInput.files[0]);
  } else {
    alert("Please upload an Excel file.");
  }

  submitButton.disabled = false;
  submitButton.textContent = "Generate Multiple Certificates";
  multipleCertificateForm.reset();
});

const handleSingleName = async (name) => {
  const capitalizedName = capitalize(name);
  const pdfBlob = await generateCertificate(capitalizedName);
  saveAs(pdfBlob, `${capitalizedName}_e-Certificate_of_Completion.pdf`);
};

const handleExcelFile = async (file) => {
  const reader = new FileReader();
  reader.onload = async (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const names = XLSX.utils.sheet_to_json(worksheet, { header: 'A' }).map((row) => row.A);

    if (names.length === 0) {
      alert("No names found in the Excel file. Please check the file and try again.");
      return;
    }

    const zip = new JSZip();
    for (const name of names) {
      if (name && typeof name === 'string') {
        const capitalizedName = capitalize(name.trim());
        const pdfBlob = await generateCertificate(capitalizedName);
        zip.file(`${capitalizedName}_e-Certificate_of_Completion.pdf`, pdfBlob);
      }
    }

    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, 'certificates.zip');
  };
  reader.readAsArrayBuffer(file);
};

const generateCertificate = async (name) => {
  const certificateId = generateCertificateId();

  const { PDFDocument, rgb } = PDFLib;

  const exBytes = await fetch("./assets/template/Cerficate-Template.pdf").then((res) => res.arrayBuffer());
  const exFont = await fetch('./assets/fonts/GreatVibes-Regular.ttf').then((res) => res.arrayBuffer());

  const pdfDoc = await PDFDocument.load(exBytes);
  
  pdfDoc.registerFontkit(fontkit);
  const myFont = await pdfDoc.embedFont(exFont);

  const pages = pdfDoc.getPages();
  const firstP = pages[0];

  const textSize = 56;
  const textWidth = myFont.widthOfTextAtSize(name, textSize);

  firstP.drawText(name, {
    x: firstP.getWidth() / 2 - textWidth / 2,
    y: 300,
    size: textSize,
    font: myFont,
    color: rgb(0, 0, 0),
  });

  const certificateIdFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
  firstP.drawText(`Certificate ID: ${certificateId}`, {
    x: 130,
    y: 60,
    size: 10,
    font: certificateIdFont,
    color: rgb(0.5, 0.5, 0.5),
  });

  const pdfBytes = await pdfDoc.save();
  return new Blob([pdfBytes], { type: 'application/pdf' });
};

const generateCertificateId = () => {
  return 'AXM' + Math.random().toString(36).substr(2, 20);
};
