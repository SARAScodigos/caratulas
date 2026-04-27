const coverTemplates = {
  utp: {
    university: "Universidad Tecnológica del Perú",
    faculty: "Facultad de Ingeniería",
    logoPath: "images/logotipo-utp.png",
    filePrefix: "caratula-utps",
    widthLogo: 300,
    heightLogo: 75,
    logoPosition: "aboveTitle",
    universitySize: 28,
    facultySize: 24,
  },
  upao: {
    university: "Universidad Privada Antenor Orrego",
    faculty: "Facultad de Ciencias Económicas",
    logoPath: "images/logotipo-upao.png",
    filePrefix: "caratula-upao",
    logoPosition: "belowTitle",
    widthLogo: 200,
    heightLogo: 250,
    universitySize: 36,
    facultySize: 32,
  },
};

const form = document.querySelector("#coverForm");
const downloadButton = document.querySelector("#downloadButton");

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  downloadButton.disabled = true;
  downloadButton.textContent = "Generando...";

  try {
    const data = getFormData();
    const doc = await createCoverDocument(data.university, data);
    const blob = await docx.Packer.toBlob(doc);
    const univ = await document.querySelector("#universidad").value;
    if (univ == "utp") {
      downloadBlob(
        blob,
        `${coverTemplates.utp.filePrefix}-${slugify(data.topic)}.docx`,
      );
    }
    if (univ == "upao") {
      downloadBlob(
        blob,
        `${coverTemplates.upao.filePrefix}-${slugify(data.topic)}.docx`,
      );
    }
  } catch (error) {
    console.error(error);
    alert("No se pudo generar la caratula. Revisa la consola del navegador.");
  } finally {
    downloadButton.disabled = false;
    downloadButton.textContent = "Descargar DOCX";
  }
});

function getFormData() {
  return {
    course: document.querySelector("#course").value.trim(),
    topic: document.querySelector("#topic").value.trim(),
    student: document.querySelector("#student").value.trim(),
    teacher: document.querySelector("#teacher").value.trim(),
    career: document.querySelector("#career").value.trim(),
    university: document.querySelector("#universidad").value,
    city: document.querySelector("#city").value.trim(),
    year: new Date().getFullYear().toString(),
  };
}

function convertCm(cm) {
  return Math.round((cm / 2.54) * 1440);
}

async function createCoverDocument(templateName, data) {
  const template = coverTemplates[templateName];
  const logo = await loadImage(template.logoPath);
  const logoElement = centeredImage(
    logo,
    template.widthLogo,
    template.heightLogo,
  );

  const titleElements = [
    centeredText(template.university, template.universitySize, true),
    centeredText(template.faculty, template.facultySize, true),
  ];

  const headerChildren =
    template.logoPosition === "aboveTitle"
      ? [logoElement, spacer(500), ...titleElements]
      : [...titleElements, spacer(300), logoElement];
  return new docx.Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertCm(2.54),
              right: convertCm(2.54),
              bottom: convertCm(2.54),
              left: convertCm(2.54),
            },
          },
        },
        children: [
          ...headerChildren,
          spacer(400),
          centeredText(data.topic, 28, true),
          spacer(400),
          labelValue("Curso", data.course),

          spacer(200),
          labelValue("Estudiante", data.student),
          spacer(200),
          labelValue("Docente", data.teacher),
          spacer(200),
          labelValue("Carrera", data.career),
          spacer(700),
          centeredText(`${data.city} - Perú`, 24, false),
          centeredText(data.year, 24, false),
        ],
      },
    ],
  });
}

async function loadImage(path) {
  const response = await fetch(path);

  if (!response.ok) {
    throw new Error(`No se pudo cargar la imagen: ${path}`);
  }

  return response.arrayBuffer();
}

function centeredImage(data, width, height) {
  return new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,

    children: [
      new docx.ImageRun({
        data,
        transformation: {
          width,
          height,
        },
      }),
    ],
  });
}

function centeredText(text, size, bold) {
  return new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 160,
    },
    children: [
      new docx.TextRun({
        text,
        bold,
        size,
        font: "Arial",
      }),
    ],
  });
}

function labelValue(label, value) {
  return new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: {
      after: 220,
    },
    children: [
      new docx.TextRun({
        text: `${label}: `,
        bold: true,
        size: 24,
        font: "Arial",
      }),
      new docx.TextRun({
        text: value,
        size: 24,
        font: "Arial",
      }),
    ],
  });
}

function spacer(after) {
  return new docx.Paragraph({
    spacing: {
      after,
    },
  });
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  link.click();

  URL.revokeObjectURL(url);
}

function slugify(value) {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}
