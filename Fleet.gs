function Fleet() {
  // The starting row location of attributes in the sheet
  this.startRow = 6;
}

/**
 * Write fleet member information to the spreadsheet. Use member counter to locate insert point
 * @param {JSON} memberInfo
 * @param {Integer} counter
 *
 * @return {boolean}
 */
Fleet.prototype.insertMembers = function(members, squads) {
  var squadInfo = [];

  for each ( squad in squads ) {
    squadInfo.push( { 'id': squad.id, 'name': squad.name } );
  }

  var insertRow = this.startRow;
  for each ( member in members ) {
    if ( member.display_name ) {
      fleetMembersSheet.getRange( "B"+insertRow ).setValue( member.display_name );
      fleetMembersSheet.getRange( "D"+insertRow ).setValue( member.level );
      fleetMembersSheet.getRange( "F"+insertRow ).setValue( member.rank );
      fleetMembersSheet.getRange( "H"+insertRow ).setValue( this.getSquadName_(member.squad_id, squadInfo) );
      fleetMembersSheet.getRange( "J"+insertRow ).setValue( member.squad_rank );
      fleetMembersSheet.getRange( "L"+insertRow ).setValue( member.event_rank );
      // fleetMembersSheet.getRange( "N"+insertRow ).setValue( member.starbase_activity );
      fleetMembersSheet.getRange( "P"+insertRow ).setValue( member.daily_activity );
      fleetMembersSheet.getRange( "R"+insertRow ).setValue( this.duration_(member.last_active) );
      insertRow++
    }
  }
}


/**
 * Clear the working area of the sheet
 */
Fleet.prototype.clear = function() {
  fleetMembersSheet.getRange("B6:R57").clear();
}

/**
 * Look up the squad name from the id
 * @return {String}
 */
Fleet.prototype.getSquadName_ = function(id, squadInfo) {
  for each( squad in squadInfo) {
    if ( squad.id == id ) { return squad.name }
  }
}

Fleet.prototype.duration_ = function(seconds) {
  var hours = Math.floor(seconds / (60 * 60));
  var minutes = Math.floor(seconds % (60 * 60) / 60);
  return hours + ":" + minutes + ":00";
}
