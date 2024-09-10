const generateExcelBtn = document.querySelector(".generateExcelBtn");

const billRates = {
  "Functional Lead (US)": 98.8,
  "Tech Lead (India)": 23.32,
  "Developer (India)": 22.26,
  "Compliance Lead (India)": 26.5,
  "Tester (India)": 12.5,
  "Test Lead (India)": 42.25,
  "Scrum Master (India)": 33.75,
};

const generateExcel = () => {
  // Get selected roles by checking which checkboxes are checked
  const selectedRoles = Array.from(
    document.querySelectorAll(".roles:checked")
  ).map((role) => role.value);
  const totalWeeks = parseInt(document.querySelector(".totalWeeks").value) || 1;

  const scope = document.querySelector(".scope").value;
  const outOfScope = document.querySelector(".outOfScope").value;
  const assumptions = document.querySelector(".assumptions").value;

  const weekColumns = [];
  for (let i = 1; i <= totalWeeks; i++) {
    weekColumns.push(`Wk${i}`);
  }

  const headers = [
    "RL",
    ...weekColumns,
    "Total Hrs",
    "Total PD",
    "Bill Rate (USD)",
    "Total Cost",
  ];

  const roleData = selectedRoles.map((role) => {
    const totalHrs = 0;
    const totalPd = 0;
    const billRate = billRates[role];
    const totalCost = totalHrs * billRate;

    return [
      role,
      ...Array(totalWeeks).fill(0),
      totalHrs,
      totalPd,
      billRate,
      totalCost,
    ];
  });

  const excelData = [
    ["Scope", scope],
    ["Out of Scope", outOfScope],
    ["Assumptions", assumptions],
    [],
    headers,
    ...roleData,
  ];

  // Create a new workbook and a worksheet from the data
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(excelData);

  // Append the worksheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, "Project Plan");

  // Generate the Excel file and trigger the download
  XLSX.writeFile(wb, "ProjectPlan.xlsx");
};

generateExcelBtn.addEventListener("click", generateExcel);
