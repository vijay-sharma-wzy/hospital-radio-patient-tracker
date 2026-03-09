/*** WARD FORM FUNCTIONS ***/


/**
 * getWardFormData
 * Consolidates all inputs from ward form into an object
 *
 * @param {Sheet} wardSheet - Ward form reference
 * @param {Object} cellPos - Stores all cell positions, incl. ranges, for form input elements
 *    (e.g. {  wardVisitsNonBool: 'H9:L38', wardVisitsBool: 'M9:M38', ward2RoomVisits: 'E11'  ...})
 * @return {Object} Ward form inputs
 */
function getWardFormData(wardSheet, cellPos) {

  // Get ward visits table as 2D array
  const wardVisits = getWardVisits(wardSheet, cellPos.wardVisits);

  return  {
    wardVisits: wardVisits,
    wardVisitsStr: getWardVisitsStr(wardVisits),
    uniqueWardsVisited: getUniqueWardsVisited(wardVisits)  
  };
}


/**
 * getWardVisits
 * Get cleaned ward visits table as 2D array
 *
 * @param {Sheet} wardSheet - Ward form reference
 * @param {string} cellName - 2D Cell range associated with ward visits table (e.g. 'I9:M31')
 * @return {Array<Array<string|boolean>>} Ward visits table as 2D array, cleaned
 */
function getWardVisits(wardSheet, cellName) {
  // Get raw table as 2D array
  let wardVisits = wardSheet.getRange(cellName).getValues();

  // Convert all columns to string except last 1 cols (bool value)
  for (let i = 0; i < wardVisits.length; i++) {
    for (let j = 0; j < wardVisits[i].length - 1; j++) {
      wardVisits[i][j] = (wardVisits[i][j] || "").toString().trim();
    }
  }

  return wardVisits
    //.filter(row => !(row[0] === "" || row[1] === ""))     // Keep the row only if both ward and bed are non-empty.
    .filter(row => !(row[0] === ""))                        // Keep row if ward is non-empty.
    .sort((rowA, rowB) => rowA[0].localeCompare(rowB[0]));  // Sort ward visits by ward number
}


/**
 * getUniqueWardsVisited
 * Get unique list of visited wards
 *
 * @param {Array<Array<string|boolean>>} wardVisits - Ward visits table as 2D array, cleaned
 * @return {string[]} Sorted array of unique ward numbers visited
 */
function getUniqueWardsVisited(wardVisits) {
  return [...new Set(wardVisits.map(row => row[0]))]      // De-dupe ward numbers
    .map(ward => String(ward))                            // Ensure ward numbers are all strings (e.g. to account for '4A')
    .sort((wardA, wardB) => wardA.localeCompare(wardB));  // Sort ward numbers
}


/**
 * getWardVisitsStr
 * Get a human-readable string representation of the ward visits table
 *
 * @param {Array<Array<string|boolean>>} wardVisits - Ward visits table as 2D array, cleaned
 * @return {string} String repr of ward visits table
 */
function getWardVisitsStr(wardVisits) {
  return '** WARD VISITS **\n\n' + wardVisits.map((row, i) => {
    const [ward, bed, name, age, activities, giftsGiven] = row;

    return `#${i + 1}
  Ward: ${ward || "N/A"}
  Bed: ${bed || "N/A"}
  Name: ${name || "N/A"}
  Age: ${age || "N/A"}
  Activities: ${activities || "N/A"}
  Gifts Given: ${giftsGiven ? "Yes" : "No"}`;
  })
  .filter(block => block !== "")     // drop fully empty rows
  .join("\n\n");                     // double newline between records
}


