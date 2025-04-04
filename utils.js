function isStudentRow(row) {
  return (
    row.length >= 4 &&
    !isNaN(row[0]) &&
    typeof row[1] === "string" &&
    typeof row[2] === "string" &&
    typeof row[3] === "string"
  );
}

function parseClass(classStr) {
  if (!classStr) return { grade: 0, letter: "" };
  var match = classStr
    .toString()
    .trim()
    .match(/(\d+)[-]?\s*([А-Яа-яA-Za-z]?)/);
  return {
    grade: match ? parseInt(match[1]) : 0,
    letter: match && match[2] ? match[2].toUpperCase() : "",
  };
}
