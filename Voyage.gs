function Voyage() {
  this.startRow = 11;
}


/**
 * Calculate the best voyage ship
 */
Voyage.prototype.bestVoyageShip = function(ships, voyage) {
  Logger.log("bestVoyageShip()");
  Logger.log( voyage );
  voyageSheet.getRange("D6").setValue( voyage.ship_trait );
  voyageSheet.getRange("D7").setValue( this.skill_(voyage.skills.primary_skill) );
  voyageSheet.getRange("D8").setValue( this.skill_(voyage.skills.secondary_skill) );


  var insertRow = this.startRow;
  for each ( ship in ships ) {
    if ( ship.name ) {
      voyageSheet.getRange( "B"+insertRow ).setValue( ship.name );
      voyageSheet.getRange( "D"+insertRow ).setValue( ship.antimatter );
      voyageSheet.getRange( "F"+insertRow ).setValue( ship.traits.toString() );
      insertRow++;
    }
  }
}

// Populate the voyage traits
Voyage.prototype.voyageTraits = function(voyage) {

  var insertRow = this.startRow;
  var slots = voyage.crew_slots;
  for each ( slot in slots ) {
    if ( slot.name ) {
      Logger.log( slot.name +" "+ slot.skill +" "+ slot.trait );
      voyageSheet.getRange("H"+insertRow).setValue( slot.name );
      voyageSheet.getRange("J"+insertRow).setValue( this.skill_(slot.skill) );
      voyageSheet.getRange("L"+insertRow).setValue( slot.trait );
      insertRow++;
    }
  }
}


/**
 * Clear the sheets working area
 */
Voyage.prototype.clear = function() {
  voyageSheet.getRange("D6:D8").clear();
  voyageSheet.getRange("B11:F").clear();
  voyageSheet.getRange("H11:L").clear();
}


Voyage.prototype.skill_ = function(msg) {
  if ( !msg || msg.length < 1 ) return msg;
  return msg.split("_")[0];
}


Voyage.prototype.cap_ = function(msg) {
  if (! msg || msg.length < 1 ) return msg;
  return msg.charAt(0).toUpperCase() + msg.slice(1);
}
