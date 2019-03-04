// ****************************************
// Main Variables
// ****************************************
var ss = SpreadsheetApp.getActiveSpreadsheet();
var profileSheet = ss.getSheetByName("Profile");
var fleetMembersSheet = ss.getSheetByName("Fleet Members");
var voyageSheet = ss.getSheetByName("Voyage");
var crewSheet = ss.getSheetByName("Crew");
var settingsSheet = ss.getSheetByName("Settings");

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui
  .createMenu('Auto Populate')
  .addItem('Load', 'login')
  .addItem('Reset', 'reset')
  .addToUi();

}

// Clean up after ourselves
function reset() {
  Logger.log("reset()");
  // profile
  profileSheet.getRange("G8").setValue("");
  profileSheet.getRange("G9").setValue("");

  profileSheet.getRange("D7").setValue("");
  profileSheet.getRange("D8").setValue("");
  profileSheet.getRange("D9").setValue("");

  // Crew sheet
  crewSheet.getRange("B5:B").clear();
  crewSheet.getRange("G5:G").clear();
  crewSheet.getRange("E5:E").setValue("FALSE");

  // Bonuses on settings
  var bonuses = new Bonuses();
  settingsSheet.getRange( bonuses.getClearRange() ).setValue("0%");

  // Fleet Member Info
  var fleet = new Fleet();
  fleet.clear();

  var voyage = new Voyage();
  voyage.clear();
}


// Login to the STT API and get an access token
function login() {
  var ui = SpreadsheetApp.getUi();
  var username = profileSheet.getRange("G8").getDisplayValue();
  var password = profileSheet.getRange("G9").getDisplayValue();

  if ( !username || !password ) return;

  // There are no true global variables so we need to pass the parsed JSON objects around to keep things tidy for now
  var api = new STTApi();
  if ( api.authenticate(username, password) ) {
    Logger.log("authenticated");
    var player = api.getPlayer().player; // access the player data to keep code cleaner
    var fleetMembers = api.getFleetMembers(player.fleet.id);
    var ships = player.character.ships;
    var voyage = player.character.voyage_descriptions[0];

    Logger.log(voyage);

    updateAllSheets(player, fleetMembers, ships, voyage);
  }
  return;
}


// Fetch the playerdata JSON
function updateAllSheets(player, fleetMembers, ships, voyage) {

  // test some sheet updates
  profileSheet.getRange("D7").setValue( player.character.display_name );
  profileSheet.getRange("D8").setValue( player.character.id );
  profileSheet.getRange("D9").setValue( player.character.level );

  //update_bonuses(player);
  //update_crew(player);
  //update_fleetMembers(fleetMembers);
  update_voyage(ships, voyage);

  return;
}

function update_fleetMembers (fleetMembers) {
  Logger.log("update_fleetMembers");
  var fleet = new Fleet();
  fleet.clear();

  var members = fleetMembers.members;
  var squads = fleetMembers.squads;
  fleet.insertMembers(members, squads);
}

function update_voyage(ships, voyage) {
  var v = new Voyage();
  v.clear();
  v.bestVoyageShip(ships, voyage);
  v.voyageTraits(voyage);
}


function update_bonuses(player) {
  Logger.log("update_bonuses");
  var bonuses = new Bonuses();

  var crewBuffs = player.character.crew_collection_buffs;
  var baseBuffs = player.character.starbase_buffs;

  // collect and condense the buffs from crew collection and starbase
  var nameRegEx = new RegExp("^[a-z]*_[a-z]*");
  for each (var buff in crewBuffs.concat(baseBuffs)) {
    if ( buff.stat != undefined ) {

      var buffInfo = {
        name: nameRegEx.exec(buff.stat)[0],
        type: '',
        source: buff.source,
        value: buff.value,
      }

      if ( buff.operator == "multiplier" ) { buffInfo.value = buff.value -1; } // multiplier are stored in a 1.xx format
      buffInfo.value = Math.round(buffInfo.value * 100); // convert to a percentage

      // range currently has both a range_max & range_min, but they are the same, we will just use range_max
      if ( buff.stat.search(/range_max/) > 0 ) { buffInfo.type = "range"; } // determine the type based on the name
      if ( buff.stat.search(/core/) > 0 ) { buffInfo.type = "core"; }

      if ( bonuses.isValid( buffInfo.name, buffInfo.type ) ) { // only push skills we're interested in, and ones with a type range_max or core
        var buffRange = bonuses.findBuffRange(buffInfo.name, buffInfo.source, buffInfo.type);
        settingsSheet.getRange( buffRange ).setValue( buffInfo.value + "%" );
      }
    }
  }
  return;
}


// player.character.crew[id, name, short_name, level, rarity]
function update_crew(player) {
  Logger.log("update_crew");

  // get a list of crew names to match on
  var crewColumn = crewSheet.getRange("A:A").getValues();

  Logger.log( crewColumn );

  // go over the crew data
  var crewData = player.character.crew;
  for(i in crewData) {
    var crewRow = crewColumn.findIndex(crewData[i].name)
    if ( crewRow ) {
      Logger.log( "Found crew " + crewData[i].name + " in: A" + crewRow);
      Logger.log( "  Found crew " + crewData[i].name + " in: A" + crewColumn.indexOf(crewData[i].name) );

      crewSheet.getRange("B"+crewRow).setValue( crewData[i].rarity );
      crewSheet.getRange("E"+crewRow).setValue( "TRUE" );
      crewSheet.getRange("G"+crewRow).setValue( crewData[i].level );
    }
  }
  return;
}


/**
 * Helper for finding text in a list, is there a faster way to do this?
 */
Array.prototype.findIndex = function(search){
  if(search == "") return false;
  for (var i=0; i<this.length; i++)
    if (this[i] == search) return i+1; // zero based index so add 1
  return false;
}

/**
 * Generate HTML query string from given object
 * Adapted from http://stackoverflow.com/a/18116302/1677912
 */
function toQuery_(obj) {return "?"+Object.keys(obj).reduce(function(a,k){a.push(k+'='+encodeURIComponent(obj[k]));return a},[]).join('&')};
