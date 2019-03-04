function Bonuses() {
  // The buffs we're interested in
  this.skills = ['command_skill', 'diplomacy_skill', 'security_skill', 'engineering_skill', 'science_skill', 'medicine_skill'];

  // The location of the skills rows and buff columns on the Settings spreadsheet
  this.skillsRows = ['25', '26', '27', '28', '29', '30'];
  this.buffCols = ['E','F','G','H'];
}

/**
 * Find the corresponding col:row of the buff on the spreadsheet
 * @param {String} buffName
 * @param {String} buffSource
 * @param {String} buffType
 *
 * @return {String} in a1Notation
 */
Bonuses.prototype.findBuffRange = function(buffName, buffSource, buffType) {
  var row = this.skillsRows[ this.skills.indexOf( buffName ) ];
  var colIndex = 0;
  if ( buffSource == "crew_collection" ) { colIndex = 2; }
  if ( buffType == "range" ) { colIndex++; }
  var col = this.buffCols[colIndex];
  Logger.log( buffName + " "+ col +""+ row );
  return (col +"" + row );
}


/**
 * Valid skills are in our skills list and with a type range_max or core
 * @param {String} buffName
 * @param {String} buffType
 *
 * @return {boolean}
 */
Bonuses.prototype.isValid = function(buffName, buffType) {
  if ( this.skills.indexOf(buffName) >= 0 && buffType != "" ) {
    return true;
  }
  return false;
}


/**
 * The range of col and rows that represent the buffs section of the Settings spreadsheet
 * @return {String} in a1Notation
 */
Bonuses.prototype.getClearRange = function() {
  // E25:H30
  return this.buffCols[0]+""+this.skillsRows[0]+":"+this.buffCols.slice(-1)[0]+""+this.skillsRows.slice(-1)[0];
}
