
/*** STUDIO FORM FUNCTIONS ***/


/**
 * getStudioFormData
 * Consolidates all inputs from studio form into an object
 *
 * @param {Sheet} studioSheet - Studio form reference
 * @param {Object} cellPos - Stores all cell positions, incl. ranges, for form input elements
 *    (e.g. {  showAired: 'E12', songRequests: 'O8:O31', studioVisits: 'I9:M31' ...})
 * @return {Object} Studio form inputs
 */
function getStudioFormData(studioSheet, cellPos) {

  // Get studio visits table as 2D array
  const studioVisits = getStudioVisits(studioSheet, cellPos.studioVisits);

  return {
    date: studioSheet.getRange(cellPos.date).getValue(),
    volsTeamLeads: getVolsList(studioSheet, cellPos.volsTeamLeads),
    volsDJs: getVolsList(studioSheet, cellPos.volsDJs),
    volsOnShift: getVolsList(studioSheet, cellPos.volsOnShift),
    volsMakeup: getVolsList(studioSheet, cellPos.volsMakeup),
    addNotes: studioSheet.getRange(cellPos.addNotes).getValue(),
    wrapUp: studioSheet.getRange(cellPos.wrapUp).getValue(),
    nextShiftMsg: studioSheet.getRange(cellPos.nextShiftMsg).getValue(),
    radioWorks: studioSheet.getRange(cellPos.radioWorks).getValue(),
    sustServiceWorks: studioSheet.getRange(cellPos.sustServiceWorks).getValue(),
    showAired: studioSheet.getRange(cellPos.showAired).getValue(),
    songRequests: getSongRequests(studioSheet, cellPos.songRequests),
    studioVisits: studioVisits,
    studioVisitsStr: getStudioVisitsStr(studioVisits)
  };
}


/**
 * getStudioVisits
 * Get cleaned studio visits table as 2D array
 *
 * @param {Sheet} studioSheet - Studio form reference
 * @param {string} cellName - 2D Cell range associated with studio visits table (e.g. 'I9:M31')
 * @return {Array<Array<string|boolean>>} Studio visits table as 2D array, cleaned
 */
function getStudioVisits(studioSheet, cellName) {
  // Get raw table as 2D array
  let studioVisits = studioSheet.getRange(cellName).getValues();

  // Convert all columns to string except last 2 cols (bool values)
  for (let i = 0; i < studioVisits.length; i++) {
    for (let j = 0; j < studioVisits[i].length - 2; j++) {
      studioVisits[i][j] = (studioVisits[i][j] || "").toString().trim();
    }
  }

  // Keep the row if at least one of name or age is non-empty.
  return studioVisits.filter(row => row[0] !== "" || row[1] !== "")
}


/**
 * getVolsList
 * Get array of volunteers from any generic volunteer multi-selecti input
 *
 * @param {Sheet} studioSheet - Studio form reference
 * @param {string} cellName - Cell associated with a volunteer multi-select (e.g. 'J3')
 * @return {string[]} Cleaned array of volunteer names
 */
function getVolsList(studioSheet, cellName) {
  let vols = studioSheet.getRange(cellName).getValue();
  // Trim volunteer names and drop empty strings
  vols = vols.split(",").map(v => v.trim()).filter(v => v.length > 0)
  return vols
}


/**
 * getSongRequests
 * Get array of song requests from song requests table
 *
 * @param {Sheet} studioSheet - Studio form reference
 * @param {string} cellName - 2D Cell range associated with song requests table (e.g. 'O8:O31')
 * @return {string[]} Cleaned array of song requests
 */
function getSongRequests(studioSheet, cellName) {
  // Trim song requests and drop empty strings
  return studioSheet.getRange(cellName)
    .getValues()
    .flat()
    .filter(songReq => songReq.trim() !== "")
    .map(songReq => (songReq || "").toString().trim());
}


/**
 * getStudioVisitsStr
 * Get a human-readable string representation of the studio visits table
 *
 * @param {Array<Array<string|boolean>>} studioVisits - Studio visits table as 2D array, cleaned
 * @return {string} String repr of studio visits table
 */
function getStudioVisitsStr(studioVisits) {
  return '** STUDIO VISITS **\n\n' + studioVisits.map((row, i) => {
    const [name, age, activities, giftsGiven, radio] = row;

    return `#${i + 1}
  Name: ${name || "N/A"}
  Age: ${age || "N/A"}
  Activities: ${activities || "N/A"}
  Gifts Given: ${giftsGiven ? "Yes" : "No"}
  Radio Show Appearance: ${radio ? "Yes" : "No"}`;
  })
  .filter(block => block !== "")     // Drop fully empty rows
  .join("\n\n");                     // Double newline between records
}

