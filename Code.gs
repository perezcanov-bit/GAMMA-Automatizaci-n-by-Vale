function buildMasterFromResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const responsesSheet = ss.getSheetByName("Form Responses 1");
  const masterSheet = ss.getSheetByName("MASTER");

  if (!responsesSheet) {
    throw new Error('Sheet "Form_Responses" was not found.');
  }

  if (!masterSheet) {
    throw new Error('Sheet "MASTER" was not found.');
  }

  const data = responsesSheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log("There are no registrations to process.");
    return;
  }

  const headers = data[0];
  const output = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i].join("").trim() === "") {
      continue;
    }

    const masterRow = createMasterRow(headers, data[i]);
    output.push(masterRow);
  }

  // Clear old MASTER data but keep row 1 headers
  if (masterSheet.getLastRow() > 1) {
    masterSheet
      .getRange(2, 1, masterSheet.getLastRow() - 1, 27)
      .clearContent();
  }

  if (output.length > 0) {
    masterSheet
      .getRange(2, 1, output.length, 27)
      .setValues(output);
  }

  Logger.log(output.length + " registrations copied to MASTER.");
}


function onRegistrationSubmit(e) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("MASTER");

  if (!masterSheet) {
    throw new Error('Sheet "MASTER" was not found.');
  }

  const responseSheet = e.range.getSheet();

  // Only run for the registration Form
  if (responseSheet.getName() !== "Form Responses 1") {
    return;
  }

  const lastColumn = responseSheet.getLastColumn();

  const headers = responseSheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  const responseRow = responseSheet
    .getRange(e.range.getRow(), 1, 1, lastColumn)
    .getValues()[0];

  const masterRow = createMasterRow(
    headers,
    responseRow
  );

  // Add new student to MASTER
  masterSheet.appendRow(masterRow);

  const newMasterRow = masterSheet.getLastRow();

  // Assign ONLY the new student
  assignNewStudent(newMasterRow);

  // Refresh all separate sheets
  updateAllGammaSheets();

  Logger.log(
    "New registration processed automatically."
  );
}

function createMasterRow(headers, row) {
  const timestamp = getResponseValue(
    headers,
    row,
    ["Timestamp"]
  );

  const firstNames = getResponseValue(
    headers,
    row,
    [
      "Nombres (del estudiante)",
      "Nombre(s) (del estudiante)",
      "Nombres del estudiante"
    ]
  );

  const lastNames = getResponseValue(
    headers,
    row,
    [
      "Apellidos (del estudiante)",
      "Apellidos del estudiante"
    ]
  );

  const fullName = (
    String(firstNames || "").trim() +
    " " +
    String(lastNames || "").trim()
  ).trim();

  const age = getResponseValue(
    headers,
    row,
    [
      "Edad (del estudiante)",
      "Edad del estudiante"
    ]
  );

  const shirtSize = getResponseValue(
    headers,
    row,
    [
      "Talla de camiseta (del estudiante)",
      "Talla de camiseta del estudiante",
      "Talla de camiseta"
    ]
  );

  const adult1 = getResponseValue(
    headers,
    row,
    [
      "Nombre y apellido del adulto responsable #1",
      "Nombre y apellido del adulto responsable 1",
      "Nombre completo del adulto responsable #1"
    ]
  );

  const phone1 = getResponseValue(
    headers,
    row,
    [
      "Número celular (del tutor responsable)",
      "Número celular del adulto responsable #1",
      "Número celular del tutor responsable #1",
      "Teléfono del adulto responsable #1"
    ]
  );

  const email1 = getResponseValue(
    headers,
    row,
    [
      "Correo electrónico del adulto responsable #1",
      "Correo electrónico del adulto responsable 1",
      "Correo del adulto responsable #1"
    ]
  );

  const adult2 = getResponseValue(
    headers,
    row,
    [
      "Nombre y apellido del adulto responsable #2",
      "Nombre y apellido del adulto responsable 2",
      "Nombre completo del adulto responsable #2"
    ]
  );

  const phone2 = getResponseValue(
    headers,
    row,
    [
      "Número celular (del tutor responsable) #2",
      "Número celular del adulto responsable #2",
      "Número celular del tutor responsable #2",
      "Teléfono del adulto responsable #2"
    ]
  );

  const email2 = getResponseValue(
    headers,
    row,
    [
      "Correo electrónico del adulto responsable #2",
      "Correo electrónico del adulto responsable 2",
      "Correo del adulto responsable #2"
    ]
  );

  const adult3 = getResponseValue(
    headers,
    row,
    [
      "Nombre y apellido del adulto responsable #3",
      "Nombre y apellido del adulto responsable 3",
      "Nombre completo del adulto responsable #3"
    ]
  );

  const phone3 = getResponseValue(
    headers,
    row,
    [
      "Número celular (del tutor responsable) #3",
      "Número celular del adulto responsable #3",
      "Número celular del tutor responsable #3",
      "Teléfono del adulto responsable #3"
    ]
  );

  const adult4 = getResponseValue(
    headers,
    row,
    [
      "Nombre y apellido del adulto responsable #4",
      "Nombre y apellido del adulto responsable 4",
      "Nombre completo del adulto responsable #4"
    ]
  );

  const phone4 = getResponseValue(
    headers,
    row,
    [
      "Número celular (del tutor responsable) #4",
      "Número celular del adulto responsable #4",
      "Número celular del tutor responsable #4",
      "Teléfono del adulto responsable #4"
    ]
  );

  const medicalCondition = getResponseValue(
    headers,
    row,
    [
      "¿Padece el estudiante algún tipo de enfermedad o problema de salud? No/Sí ¿Cuál?",
      "¿Padece el estudiante algún tipo de enfermedad o problema de salud?",
      "Enfermedades"
    ]
  );

  const allergy = getResponseValue(
    headers,
    row,
    [
      "¿Tiene el estudiante algún tipo de alergia? Especifique.",
      "¿Tiene el estudiante algún tipo de alergia?",
      "Alergias"
    ]
  );

  const otherCondition = getResponseValue(
    headers,
    row,
    [
      "¿Hay alguna condición del estudiante que usted considere que deba ser de nuestro conocimiento? No/Sí ¿Cuál?",
      "¿Hay alguna condición del estudiante que usted considere que deba ser de nuestro conocimiento?",
      "Otra condición"
    ]
  );

  const insurance = getResponseValue(
    headers,
    row,
    [
      "¿Cuenta con algún seguro médico?",
      "Seguro médico"
    ]
  );

  const optionA = getResponseValue(
    headers,
    row,
    [
      "Clase periodo 1",
      "Clase período 1",
      "Clase Periodo 1"
    ]
  );

  const optionB = getResponseValue(
    headers,
    row,
    [
      "Clase periodo 2",
      "Clase período 2",
      "Clase Periodo 2"
    ]
  );

  const optionC = getResponseValue(
    headers,
    row,
    [
      "Alternativa",
      "Clase alternativa",
      "Alternativa C"
    ]
  );

  return [
    timestamp,         // A  Fecha de inscripción
    fullName,          // B  Nombre completo
    age,               // C  Edad
    shirtSize,         // D  Talla de camiseta

    adult1,            // E
    phone1,            // F
    email1,            // G

    adult2,            // H
    phone2,            // I
    email2,            // J

    adult3,            // K
    phone3,            // L

    adult4,            // M
    phone4,            // N

    medicalCondition,  // O
    allergy,           // P
    otherCondition,    // Q
    insurance,         // R

    optionA,           // S
    optionB,           // T
    optionC,           // U

    "",                // V Asignación automática P1
    "",                // W Asignación automática P2
    "",                // X Asignación final P1
    "",                // Y Asignación final P2
    "",                // Z Estado
    ""                 // AA Alerta
  ];
}


function getResponseValue(headers, row, possibleNames) {
  for (let i = 0; i < possibleNames.length; i++) {
    const targetHeader = normalizeHeader(possibleNames[i]);

    for (let j = 0; j < headers.length; j++) {
      const currentHeader = normalizeHeader(headers[j]);

      if (currentHeader === targetHeader) {
        return row[j];
      }
    }
  }

  return "";
}


function normalizeHeader(text) {
  return String(text || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}
const MAX_CLASS_SIZE = 20;

const NORMAL_CLASSES = [
  "Inglés A: Principiante",
  "Inglés B: Intermedio/Avanzado",
  "Lengua (Español)",
  "Matemáticas A: Principiante",
  "Matemáticas B: Estudiantes con conocimiento de multiplicación",
  "Computación",
  "Laboratorio",
  "Arte",
  "Tutorias"
];


// ======================================================
// ASSIGN ONLY ONE NEW STUDENT
// ======================================================

function assignNewStudent(rowNumber) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("MASTER");

  if (!master) {
    throw new Error('Sheet "MASTER" was not found.');
  }

  const student = master
    .getRange(rowNumber, 1, 1, 27)
    .getValues()[0];

  const age = Number(student[2]);

  const optionA = student[18];
  const optionB = student[19];
  const optionC = student[20];

  const period1Counts = createEmptyCounts();
  const period2Counts = createEmptyCounts();

  // Count all previous students using FINAL assignments
  if (rowNumber > 2) {

    const existingStudents = master
      .getRange(
        2,
        1,
        rowNumber - 2,
        27
      )
      .getValues();

    existingStudents.forEach(row => {

      incrementClassCount(
        period1Counts,
        row[23]
      );

      incrementClassCount(
        period2Counts,
        row[24]
      );
    });
  }

  let autoP1 = "";
  let autoP2 = "";
  let status = "";
  let alert = "";


  // ======================================================
  // KINDER RULES
  // ======================================================

  const kinderFormOption =
    "Kinder: Niños y niñas entre 4 a 6 años de edad.";

  const choseKinder =
    normalizeClassName(optionA) ===
      normalizeClassName(kinderFormOption) ||
    normalizeClassName(optionB) ===
      normalizeClassName(kinderFormOption) ||
    normalizeClassName(optionC) ===
      normalizeClassName(kinderFormOption);


  // AGE 4
  if (age === 4) {

    if (
      getClassCount(period1Counts, "Kinder 1") <
        MAX_CLASS_SIZE &&
      getClassCount(period2Counts, "Kinder 1") <
        MAX_CLASS_SIZE
    ) {

      autoP1 = "Kinder 1";
      autoP2 = "Kinder 1";

      status = "KINDER";
      alert = "KINDER OVERRIDE - CONTACT PARENT";

    } else {

      status = "MANUAL REVIEW";
      alert = "KINDER 1 FULL";
    }
  }


  // AGE 5
  else if (age === 5) {

    const kinder1Available =
      getClassCount(period1Counts, "Kinder 1") <
        MAX_CLASS_SIZE &&
      getClassCount(period2Counts, "Kinder 1") <
        MAX_CLASS_SIZE;

    const kinder2Available =
      getClassCount(period1Counts, "Kinder 2") <
        MAX_CLASS_SIZE &&
      getClassCount(period2Counts, "Kinder 2") <
        MAX_CLASS_SIZE;

    const kinder1Total =
      getClassCount(period1Counts, "Kinder 1") +
      getClassCount(period2Counts, "Kinder 1");

    const kinder2Total =
      getClassCount(period1Counts, "Kinder 2") +
      getClassCount(period2Counts, "Kinder 2");


    if (
      kinder1Available &&
      (
        !kinder2Available ||
        kinder1Total <= kinder2Total
      )
    ) {

      autoP1 = "Kinder 1";
      autoP2 = "Kinder 1";

      status = "KINDER";
      alert = "KINDER OVERRIDE - CONTACT PARENT";

    }

    else if (kinder2Available) {

      autoP1 = "Kinder 2";
      autoP2 = "Kinder 2";

      status = "KINDER";
      alert = "KINDER OVERRIDE - CONTACT PARENT";

    }

    else {

      status = "MANUAL REVIEW";
      alert = "KINDER CLASSES FULL";
    }
  }


  // AGE 6 WHO CHOSE KINDER
  else if (age === 6 && choseKinder) {

    if (
      getClassCount(period1Counts, "Kinder 2") <
        MAX_CLASS_SIZE &&
      getClassCount(period2Counts, "Kinder 2") <
        MAX_CLASS_SIZE
    ) {

      autoP1 = "Kinder 2";
      autoP2 = "Kinder 2";

      status = "KINDER";
      alert = "";

    } else {

      status = "MANUAL REVIEW";
      alert = "KINDER 2 FULL";
    }
  }


  // ======================================================
  // NORMAL STUDENTS
  // ======================================================

  else {

    let assigned = false;

    // First try TWO of their requested classes
    const combinations = [
      [optionA, optionB],
      [optionB, optionA],
      [optionA, optionC],
      [optionC, optionB],
      [optionC, optionA],
      [optionB, optionC]
    ];


    for (
      let i = 0;
      i < combinations.length;
      i++
    ) {

      const p1 = combinations[i][0];
      const p2 = combinations[i][1];

      if (
        isValidNormalClass(p1) &&
        isValidNormalClass(p2) &&
        normalizeClassName(p1) !==
          normalizeClassName(p2) &&
        getClassCount(period1Counts, p1) <
          MAX_CLASS_SIZE &&
        getClassCount(period2Counts, p2) <
          MAX_CLASS_SIZE
      ) {

        autoP1 = getCanonicalClassName(p1);
        autoP2 = getCanonicalClassName(p2);

        status = "AUTO";
        assigned = true;

        break;
      }
    }


    // ======================================================
    // KEEP AT LEAST ONE REQUESTED CLASS
    // ======================================================

    if (!assigned) {

      const requestedPlacements = [

        {
          className: optionA,
          period: 1
        },

        {
          className: optionB,
          period: 2
        },

        {
          className: optionC,
          period: 1
        },

        {
          className: optionC,
          period: 2
        },

        {
          className: optionA,
          period: 2
        },

        {
          className: optionB,
          period: 1
        }
      ];


      for (
        let i = 0;
        i < requestedPlacements.length &&
        !assigned;
        i++
      ) {

        const requested =
          requestedPlacements[i].className;

        const period =
          requestedPlacements[i].period;

        if (!isValidNormalClass(requested)) {
          continue;
        }

        const requestedClass =
          getCanonicalClassName(requested);


        if (
          period === 1 &&
          getClassCount(
            period1Counts,
            requestedClass
          ) < MAX_CLASS_SIZE
        ) {

          const fallbackP2 =
            getLeastFullAvailableClass(
              period2Counts,
              requestedClass
            );

          if (fallbackP2) {

            autoP1 = requestedClass;
            autoP2 = fallbackP2;

            assigned = true;
          }
        }


        else if (
          period === 2 &&
          getClassCount(
            period2Counts,
            requestedClass
          ) < MAX_CLASS_SIZE
        ) {

          const fallbackP1 =
            getLeastFullAvailableClass(
              period1Counts,
              requestedClass
            );

          if (fallbackP1) {

            autoP1 = fallbackP1;
            autoP2 = requestedClass;

            assigned = true;
          }
        }
      }


      if (assigned) {

        status = "PARTIAL CHOICE";

        alert =
          "ONE CLASS OUTSIDE CHOICES - CONTACT PARENT";
      }
    }


    // ======================================================
    // ABSOLUTE LAST RESORT
    // ======================================================

    if (!assigned) {

      autoP1 =
        getLeastFullAvailableClass(
          period1Counts
        );

      autoP2 =
        getLeastFullAvailableClass(
          period2Counts,
          autoP1
        );

      if (autoP1 && autoP2) {

        status = "OUTSIDE CHOICES";

        alert =
          "NO REQUESTED CLASSES AVAILABLE - CONTACT PARENT";

      } else {

        status = "MANUAL REVIEW";

        alert =
          "NO AVAILABLE CLASS COMBINATION";
      }
    }
  }


  // ======================================================
  // WRITE RESULT TO MASTER
  // ======================================================

  master
    .getRange(
      rowNumber,
      22,
      1,
      6
    )
    .setValues([[
      autoP1,
      autoP2,
      autoP1,
      autoP2,
      status,
      alert
    ]]);


  Logger.log(
    "New student assigned: " +
    autoP1 +
    " / " +
    autoP2
  );
}


// ======================================================
// HELPERS
// ======================================================

function createEmptyCounts() {

  const counts = {};

  NORMAL_CLASSES.forEach(className => {
    counts[className] = 0;
  });

  counts["Kinder 1"] = 0;
  counts["Kinder 2"] = 0;

  return counts;
}


function incrementClassCount(
  counts,
  className
) {

  const text =
    String(className || "").trim();

  if (!text) {
    return;
  }

  const normalized =
    normalizeClassName(text);


  if (
    normalized ===
    normalizeClassName("Kinder 1")
  ) {

    counts["Kinder 1"]++;
    return;
  }


  if (
    normalized ===
    normalizeClassName("Kinder 2")
  ) {

    counts["Kinder 2"]++;
    return;
  }


  const canonical =
    getCanonicalClassName(text);

  if (canonical) {
    counts[canonical]++;
  }
}


function isValidNormalClass(className) {

  return (
    getCanonicalClassName(className) !== ""
  );
}


function getCanonicalClassName(className) {

  const normalized =
    normalizeClassName(className);

  for (
    let i = 0;
    i < NORMAL_CLASSES.length;
    i++
  ) {

    if (
      normalizeClassName(
        NORMAL_CLASSES[i]
      ) === normalized
    ) {

      return NORMAL_CLASSES[i];
    }
  }

  return "";
}


function getClassCount(
  counts,
  className
) {

  const text =
    String(className || "").trim();

  if (!text) {
    return 0;
  }

  const normalized =
    normalizeClassName(text);


  if (
    normalized ===
    normalizeClassName("Kinder 1")
  ) {

    return counts["Kinder 1"] || 0;
  }


  if (
    normalized ===
    normalizeClassName("Kinder 2")
  ) {

    return counts["Kinder 2"] || 0;
  }


  const canonical =
    getCanonicalClassName(text);

  if (!canonical) {
    return 0;
  }

  return counts[canonical] || 0;
}


function getLeastFullAvailableClass(
  counts,
  excludedClass = ""
) {

  const excluded =
    normalizeClassName(excludedClass);

  let bestClass = "";
  let smallestCount = Infinity;


  NORMAL_CLASSES.forEach(className => {

    if (
      normalizeClassName(className) ===
      excluded
    ) {
      return;
    }

    const count =
      getClassCount(
        counts,
        className
      );

    if (
      count < MAX_CLASS_SIZE &&
      count < smallestCount
    ) {

      smallestCount = count;
      bestClass = className;
    }
  });


  return bestClass;
}
const OUTPUT_SHEETS = {
  CLASS_ROSTERS: "PASTE_CLASS_ROSTERS_SPREADSHEET_ID_HERE",
  EXIT_GROUPS: "PASTE_EXIT_GROUPS_SPREADSHEET_ID_HERE",
  MEDICAL: "PASTE_MEDICAL_SPREADSHEET_ID_HERE",
  ATTENDANCE: "PASTE_ATTENDANCE_SPREADSHEET_ID_HERE"
};
function updateAllGammaSheets() {
  updateClassRosters();
  updateExitGroups();
  updateMedicalSheet();
  updateAttendanceSheets();

  Logger.log("All Gamma sheets updated successfully.");
}
function updateClassRosters() {
  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  const master = sourceSS.getSheetByName("MASTER");

  const outputSS = SpreadsheetApp.openById(
    OUTPUT_SHEETS.CLASS_ROSTERS
  );

  const data = master.getDataRange().getValues();

  if (data.length < 2) return;

  const students = data.slice(1);

  buildPeriodRoster(outputSS, "PERIODO 1", students, 23);
  buildPeriodRoster(outputSS, "PERIODO 2", students, 24);
}

function buildPeriodRoster(outputSS, sheetName, students, assignmentIndex) {

  let sheet = outputSS.getSheetByName(sheetName);

  if (!sheet) {
    sheet = outputSS.insertSheet(sheetName);
  }

  sheet.clear();

  const classes = getAllFinalClasses(students, assignmentIndex);

  let currentRow = 1;

  classes.forEach(className => {

    sheet.getRange(currentRow, 1).setValue(className);
    sheet.getRange(currentRow, 1, 1, 5).merge();

    currentRow++;

    sheet.getRange(currentRow, 1, 1, 5).setValues([[
      "Nombre",
      "Edad",
      "Talla",
      "Adulto responsable",
      "Teléfono"
    ]]);

    currentRow++;

    const classStudents = students
      .filter(row => row[assignmentIndex] === className)
      .sort((a, b) =>
        String(a[1]).localeCompare(String(b[1]), "es")
      );

    const roster = classStudents.map(row => [
      row[1],
      row[2],
      row[3],
      row[4],
      row[5]
    ]);

    if (roster.length > 0) {
      sheet
        .getRange(currentRow, 1, roster.length, 5)
        .setValues(roster);

      currentRow += roster.length;
    }

    currentRow += 2;
  });

  sheet.autoResizeColumns(1, 5);
}

function getAllFinalClasses(students, assignmentIndex) {

  const classes = students
    .map(row => String(row[assignmentIndex] || "").trim())
    .filter(value => value !== "");

  return [...new Set(classes)].sort((a, b) =>
    a.localeCompare(b, "es")
  );
}
function updateExitGroups() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // IMPORTANT:
  // Change this only if your Form response tab has a different name.
  const responsesSheet = ss.getSheetByName("Form Responses 1");

  if (!responsesSheet) {
    throw new Error('Form response sheet was not found.');
  }

  const outputSS = SpreadsheetApp.openById(
    OUTPUT_SHEETS.EXIT_GROUPS
  );

  let sheet = outputSS.getSheetByName("GRUPOS DE SALIDA");

  if (!sheet) {
    sheet = outputSS.insertSheet("GRUPOS DE SALIDA");
  }

  const data = responsesSheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log("No registrations found.");
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  const students = rows
    .filter(row => row.join("").trim() !== "")
    .map(row => {

      const firstNames = getResponseValue(headers, row, [
        "Nombres (del estudiante)",
        "Nombre(s) (del estudiante)",
        "Nombres del estudiante"
      ]);

      const lastNames = getResponseValue(headers, row, [
        "Apellidos (del estudiante)",
        "Apellidos del estudiante"
      ]);

      const tutors = [];

      // Tutor 1
      addTutorIfPresent(
        tutors,
        getResponseValue(headers, row, [
          "Nombre y apellido del adulto responsable #1",
          "Nombre y apellido del adulto responsable 1"
        ]),
        getResponseValue(headers, row, [
          "Número celular (del tutor responsable)",
          "Número celular del adulto responsable #1",
          "Número celular del tutor responsable #1"
        ]),
        getResponseValue(headers, row, [
          "Correo electrónico del adulto responsable #1",
          "Correo electrónico del adulto responsable 1",
          "Correo del adulto responsable #1"
        ])
      );

      // Tutor 2
      addTutorIfPresent(
        tutors,
        getResponseValue(headers, row, [
          "Nombre y apellido del adulto responsable #2",
          "Nombre y apellido del adulto responsable 2"
        ]),
        getResponseValue(headers, row, [
          "Número celular (del tutor responsable) #2",
          "Número celular del adulto responsable #2",
          "Número celular del tutor responsable #2"
        ]),
        getResponseValue(headers, row, [
          "Correo electrónico del adulto responsable #2",
          "Correo electrónico del adulto responsable 2",
          "Correo del adulto responsable #2"
        ])
      );

      // Tutor 3
      addTutorIfPresent(
        tutors,
        getResponseValue(headers, row, [
          "Nombre y apellido del adulto responsable #3",
          "Nombre y apellido del adulto responsable 3"
        ]),
        getResponseValue(headers, row, [
          "Número celular (del tutor responsable) #3",
          "Número celular del adulto responsable #3",
          "Número celular del tutor responsable #3"
        ]),
        ""
      );

      // Tutor 4
      addTutorIfPresent(
        tutors,
        getResponseValue(headers, row, [
          "Nombre y apellido del adulto responsable #4",
          "Nombre y apellido del adulto responsable 4"
        ]),
        getResponseValue(headers, row, [
          "Número celular (del tutor responsable) #4",
          "Número celular del adulto responsable #4",
          "Número celular del tutor responsable #4"
        ]),
        ""
      );

      return {
        firstNames: String(firstNames || "").trim(),
        lastNames: String(lastNames || "").trim(),
        tutors: tutors
      };
    });

  // Sort alphabetically by LAST NAME
  students.sort((a, b) =>
    a.lastNames.localeCompare(
      b.lastNames,
      "es",
      { sensitivity: "base" }
    )
  );

  // Remove old merged cells and old content
  sheet.getDataRange().breakApart();
  sheet.clear();

  // Headers
  sheet.getRange(1, 1, 1, 5).setValues([[
    "NOMBRE",
    "APELLIDOS",
    "TUTOR",
    "CELULAR",
    "CORREO"
  ]]);

  let currentRow = 2;

  students.forEach(student => {

    let tutors = student.tutors;

    // If there are no tutors, still create one row
    if (tutors.length === 0) {
      tutors = [{
        name: "",
        phone: "",
        email: ""
      }];
    }

    const startRow = currentRow;

    // Write tutor rows
    tutors.forEach(tutor => {

      sheet.getRange(currentRow, 3).setValue(tutor.name);
      sheet.getRange(currentRow, 4).setValue(tutor.phone);
      sheet.getRange(currentRow, 5).setValue(tutor.email);

      currentRow++;
    });

    const endRow = currentRow - 1;

    // Merge child name cells vertically
    if (endRow > startRow) {

      sheet
        .getRange(
          startRow,
          1,
          endRow - startRow + 1,
          1
        )
        .merge();

      sheet
        .getRange(
          startRow,
          2,
          endRow - startRow + 1,
          1
        )
        .merge();
    }

    // Write child names
    sheet
      .getRange(startRow, 1)
      .setValue(student.firstNames);

    sheet
      .getRange(startRow, 2)
      .setValue(student.lastNames);
  });

  // Formatting
  sheet.setFrozenRows(1);

  const lastOutputRow = Math.max(currentRow - 1, 1);

  sheet
    .getRange(1, 1, lastOutputRow, 5)
    .setVerticalAlignment("middle");

  sheet
    .getRange(1, 1, lastOutputRow, 5)
    .setWrap(true);

  sheet.autoResizeColumns(1, 5);

  Logger.log(
    students.length +
    " students added to GRUPOS DE SALIDA."
  );
}


function addTutorIfPresent(tutors, name, phone, email) {

  const tutorName = String(name || "").trim();
  const tutorPhone = String(phone || "").trim();
  const tutorEmail = String(email || "").trim();

  if (
    tutorName !== "" ||
    tutorPhone !== "" ||
    tutorEmail !== ""
  ) {
    tutors.push({
      name: tutorName,
      phone: tutorPhone,
      email: tutorEmail
    });
  }
}
function updateMedicalSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("MASTER");

  if (!master) {
    throw new Error('Sheet "MASTER" was not found.');
  }

  const outputSS = SpreadsheetApp.openById(
    OUTPUT_SHEETS.MEDICAL
  );

  const data = master.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log("No students found in MASTER.");
    return;
  }

  const students = data.slice(1);

  buildMedicalCategorySheet(
    outputSS,
    "ENFERMEDADES",
    students,
    14
  );

  buildMedicalCategorySheet(
    outputSS,
    "ALERGIAS",
    students,
    15
  );

  buildMedicalCategorySheet(
    outputSS,
    "OTRAS CONDICIONES",
    students,
    16
  );

  Logger.log("Gamma medical sheets updated successfully.");
}


function buildMedicalCategorySheet(
  outputSS,
  sheetName,
  students,
  infoIndex
) {

  let sheet = outputSS.getSheetByName(sheetName);

  if (!sheet) {
    sheet = outputSS.insertSheet(sheetName);
  }

  const filteredStudents = students
    .filter(row =>
      hasRelevantMedicalInfo(row[infoIndex])
    )
    .sort((a, b) =>
      String(a[1]).localeCompare(
        String(b[1]),
        "es",
        { sensitivity: "base" }
      )
    );

  const output = filteredStudents.map(row => [
    row[1],          // Nombre completo
    row[2],          // Edad
    row[infoIndex],  // Medical information
    row[4],          // Adulto responsable 1
    row[5],          // Teléfono 1
    row[17]          // Seguro médico
  ]);

  sheet.getDataRange().breakApart();
  sheet.clear();

  sheet.getRange(1, 1, 1, 6).setValues([[
    "NOMBRE",
    "EDAD",
    "INFORMACIÓN",
    "ADULTO RESPONSABLE",
    "CONTACTO",
    "SEGURO MÉDICO"
  ]]);

  if (output.length > 0) {
    sheet
      .getRange(2, 1, output.length, 6)
      .setValues(output);
  }

  sheet.setFrozenRows(1);

  sheet
    .getRange(
      1,
      1,
      Math.max(output.length + 1, 1),
      6
    )
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.autoResizeColumns(1, 6);
}


function hasRelevantMedicalInfo(value) {

  const text = String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  const emptyAnswers = [
    "",
    "no",
    "ninguna",
    "ninguno",
    "ninguna.",
    "ninguno.",
    "no aplica",
    "n/a",
    "na"
  ];

  return !emptyAnswers.includes(text);
}
// ======================================================
// ATTENDANCE SYSTEM
// Safely syncs final class assignments with attendance
// without moving existing students or mixing attendance.
// ======================================================

function updateAttendanceSheets() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("MASTER");

  if (!master) {
    throw new Error('Sheet "MASTER" was not found.');
  }

  const outputSS = SpreadsheetApp.openById(
    OUTPUT_SHEETS.ATTENDANCE
  );

  const data = master.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log("No students found in MASTER.");
    return;
  }

  const students = data.slice(1);


  // ======================================================
  // MASTER CLASS NAME -> ATTENDANCE TAB NAME
  // ======================================================

  const classMap = {

    "Matemáticas A: Principiante":
      "Mate A",

    "Matemáticas B: Estudiantes con conocimiento de multiplicación":
      "Mate B",

    "Inglés A: Principiante":
      "Ingles A",

    "Inglés B: Intermedio/Avanzado":
      "Ingles B",

    "Kinder 1":
      "Kinder 1",

    "Kinder 2":
      "Kinder2",

    "Lengua (Español)":
      "Lengua",

    "Computación":
      "Computación",

    "Laboratorio":
      "Laboratorio",

    "Arte":
      "Arte",

    "Tutorias":
      "Tutorias"
  };


  // ======================================================
  // PROCESS EVERY CLASS
  // ======================================================

  Object.keys(classMap).forEach(className => {

    const tabName =
      classMap[className];

    const sheet =
      outputSS.getSheetByName(tabName);


    if (!sheet) {

      Logger.log(
        'Attendance tab not found: "' +
        tabName +
        '"'
      );

      return;
    }


    // ======================================================
    // STUDENTS WHO SHOULD CURRENTLY BE IN PERIOD 1
    // Uses X = Final Assignment P1
    // ======================================================

    const period1Students = students
      .filter(row =>
        normalizeClassName(row[23]) ===
        normalizeClassName(className)
      )
      .map(row =>
        String(row[1] || "").trim()
      )
      .filter(name => name !== "");


    // ======================================================
    // STUDENTS WHO SHOULD CURRENTLY BE IN PERIOD 2
    // Uses Y = Final Assignment P2
    // ======================================================

    const period2Students = students
      .filter(row =>
        normalizeClassName(row[24]) ===
        normalizeClassName(className)
      )
      .map(row =>
        String(row[1] || "").trim()
      )
      .filter(name => name !== "");


    // ======================================================
    // SAFETY: MAXIMUM 20 STUDENTS
    // ======================================================

    if (period1Students.length > 20) {

      throw new Error(
        className +
        " has more than 20 students in Period 1."
      );
    }


    if (period2Students.length > 20) {

      throw new Error(
        className +
        " has more than 20 students in Period 2."
      );
    }


    // ======================================================
    // PERIOD 1
    // Student rows: 5 through 24
    // ======================================================

    syncAttendanceRoster(
      sheet,
      period1Students,
      5,
      24
    );


    // ======================================================
    // PERIOD 2
    // Student rows: 29 through 48
    // ======================================================

    syncAttendanceRoster(
      sheet,
      period2Students,
      29,
      48
    );


    Logger.log(
      tabName +
      ": P1 = " +
      period1Students.length +
      ", P2 = " +
      period2Students.length
    );
  });


  Logger.log(
    "Attendance rosters safely synchronized."
  );
}


// ======================================================
// SAFE ROSTER SYNCHRONIZATION
//
// Existing students DO NOT move rows.
// New students are placed in the first empty row.
// Students who leave the class have their old row cleared.
// ======================================================

function syncAttendanceRoster(
  sheet,
  desiredStudents,
  startRow,
  endRow
) {

  const numberOfRows =
    endRow - startRow + 1;


  // ======================================================
  // GET CURRENT NAMES FROM COLUMN B
  // ======================================================

  const currentNames = sheet
    .getRange(
      startRow,
      2,
      numberOfRows,
      1
    )
    .getValues()
    .map(row =>
      String(row[0] || "").trim()
    );


  // Normalize students who SHOULD be in this class
  const desiredNormalized =
    desiredStudents.map(name =>
      normalizeStudentName(name)
    );


  // ======================================================
  // STEP 1
  // REMOVE STUDENTS WHO NO LONGER BELONG IN THIS CLASS
  // ======================================================

  for (
    let i = 0;
    i < currentNames.length;
    i++
  ) {

    const currentName =
      currentNames[i];


    // Empty row
    if (!currentName) {
      continue;
    }


    const normalizedCurrent =
      normalizeStudentName(
        currentName
      );


    const stillBelongs =
      desiredNormalized.includes(
        normalizedCurrent
      );


    // Student changed class or period
    if (!stillBelongs) {

      const actualRow =
        startRow + i;


      clearAttendanceStudentRow(
        sheet,
        actualRow
      );


      // Mark this row as available
      currentNames[i] = "";
    }
  }


  // ======================================================
  // STEP 2
  // CREATE UPDATED LIST OF STUDENTS ALREADY PRESENT
  // ======================================================

  const currentNormalized =
    currentNames.map(name =>
      normalizeStudentName(name)
    );


  // ======================================================
  // STEP 3
  // ADD STUDENTS WHO ARE MISSING
  // ======================================================

  desiredStudents.forEach(
    studentName => {

      const normalizedStudent =
        normalizeStudentName(
          studentName
        );


      // Already here:
      // DO NOTHING.
      //
      // This preserves the student's exact row
      // and all existing attendance.
      if (
        currentNormalized.includes(
          normalizedStudent
        )
      ) {
        return;
      }


      // ==================================================
      // FIND FIRST EMPTY STUDENT ROW
      // ==================================================

      const emptyIndex =
        currentNames.findIndex(
          name => name === ""
        );


      if (emptyIndex === -1) {

        throw new Error(
          'No empty attendance row available in "' +
          sheet.getName() +
          '".'
        );
      }


      const actualRow =
        startRow + emptyIndex;


      // ==================================================
      // WRITE NEW STUDENT NAME ONLY
      // ==================================================

      sheet
        .getRange(
          actualRow,
          2
        )
        .setValue(
          studentName
        );


      // Update local arrays
      currentNames[emptyIndex] =
        studentName;

      currentNormalized[emptyIndex] =
        normalizedStudent;
    }
  );
}


// ======================================================
// CLEAR OLD STUDENT ATTENDANCE
//
// B = Student name
// Z = Last attendance column
//
// ONLY B:Z gets cleared.
// Columns AA onward are NEVER touched.
// Formatting is preserved.
// ======================================================

function clearAttendanceStudentRow(
  sheet,
  rowNumber
) {

  // B = column 2
  // Z = column 26
  //
  // Number of columns from B through Z:
  // 26 - 2 + 1 = 25

  sheet
    .getRange(
      rowNumber,
      2,
      1,
      25
    )
    .clearContent();
}


// ======================================================
// NORMALIZE STUDENT NAMES
// Prevents capitalization/spacing/accent differences
// from creating accidental duplicates.
// ======================================================

function normalizeStudentName(text) {

  return String(text || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(
      /[\u0300-\u036f]/g,
      ""
    )
    .replace(
      /\s+/g,
      " "
    );
}

function normalizeClassName(text) {
  return String(text || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}
function onFinalAssignmentEdit(e) {

  const sheet = e.range.getSheet();

  // Only react to edits in MASTER
  if (sheet.getName() !== "MASTER") {
    return;
  }

  const row = e.range.getRow();
  const column = e.range.getColumn();

  // Ignore header row
  if (row < 2) {
    return;
  }

  // X = 24 = Final P1
  // Y = 25 = Final P2
  if (column !== 24 && column !== 25) {
    return;
  }

  // Refresh only sheets affected by class assignment
  updateClassRosters();
  updateAttendanceSheets();

  Logger.log(
    "Final assignment changed. Rosters and attendance updated."
  );
}