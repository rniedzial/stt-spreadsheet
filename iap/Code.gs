/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/*
    Copyright (C) 2017 IAmPicard
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Start STT', 'startSTT')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function startSTT() {
  var userProperties = PropertiesService.getUserProperties();
  var accessToken = userProperties.getProperty('accessToken');
  var name = userProperties.getProperty('name');

  // If user is already logged in, show the sidebar
  if (accessToken && name) {
    showSidebar();
  } else {
    // Otherwise, show the login dialog
    logout();
  }
}

function logout() {
  // The sidebar is already closed from client script when this gets called
  var loginDialog = HtmlService.createHtmlOutputFromFile('Login').setWidth(280).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(loginDialog, 'Login to Star Trek Timelines');
}

function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar');
  ui.userName = PropertiesService.getUserProperties().getProperty('name');
  SpreadsheetApp.getUi().showSidebar(ui.evaluate().setTitle('STT Crew loader'));
}

function clearSheets() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'Please confirm',
    'Are you sure you want to delete the \'Crew roster\' and \'Cadet missions\' sheets?',
    ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Crew roster");
    if (sheet) {
      if (ss.getSheets().length == 1) {
        ss.insertSheet('Sheet1');
      }
      ss.deleteSheet(sheet);
    }

    sheet = ss.getSheetByName("Cadet missions");
    if (sheet) {
      if (ss.getSheets().length == 1) {
        ss.insertSheet('Sheet1');
      }
      ss.deleteSheet(sheet);
    }

    sheet = ss.getSheetByName("Items");
    if (sheet) {
      if (ss.getSheets().length == 1) {
        ss.insertSheet('Sheet1');
      }
      ss.deleteSheet(sheet);
    }

    sheet = ss.getSheetByName("Ships");
    if (sheet) {
      if (ss.getSheets().length == 1) {
        ss.insertSheet('Sheet1');
      }
      ss.deleteSheet(sheet);
    }

    sheet = ss.getSheetByName("Stats");
    if (sheet) {
      if (ss.getSheets().length == 1) {
        ss.insertSheet('Sheet1');
      }
      ss.deleteSheet(sheet);
    }
  } else {
    // User clicked "No" or X in the title bar.

  }
}

function login(username, password) {
  var data = 'username=' + username + '&password=' + password + '&client_id=4fc852d7-d602-476a-a292-d243022a475d&grant_type=password';

  var options = {
    'method': 'post',
    'payload': data
  };

  var response = UrlFetchApp.fetch('https://thorium.disruptorbeam.com/oauth2/token', options);
  var result = JSON.parse(response.getContentText());

  PropertiesService.getUserProperties().setProperty('accessToken', result.access_token);

  showSidebar();

  return result;
}

var SKILLS = {
  'command_skill': 'Command',
  'science_skill': 'Science',
  'security_skill': 'Security',
  'engineering_skill': 'Engineering',
  'diplomacy_skill': 'Diplomacy',
  'medicine_skill': 'Medicine'
};

function loadCrew(loadFrozen, loadItems, loadShips, loadSnap) {
  var access_token = PropertiesService.getUserProperties().getProperty('accessToken');
  var apiDomain = 'https://stt.disruptorbeam.com/';
  var apiQueryString = '?client_api=9&access_token=' + access_token;
  var response = UrlFetchApp.fetch(apiDomain + 'player' + apiQueryString);
  var playerData = JSON.parse(response.getContentText());

  response = UrlFetchApp.fetch(apiDomain + 'character/get_avatar_crew_archetypes' + apiQueryString);
  var crewArchetypes = JSON.parse(response.getContentText());
  var immortals = [];

  var result = {
    vipLevel: playerData.player.vip_level,
    name: playerData.player.character.display_name,
    level: playerData.player.character.level,
    crewLimit: playerData.player.character.crew_limit,
    crew: undefined,
    cadetMissions: []
  };

  PropertiesService.getUserProperties().setProperty('name', result.name);

  result.crew = new Object();
  crewArchetypes.crew_avatars.forEach(function (av) {
    result.crew[av.id] = { name: av.name, short_name: av.short_name, max_rarity: av.max_rarity, traits: av.traits, traits_hidden: av.traits_hidden, have: false, airlocked: false, immortal: 0 };
  });

  function findCrewById(id, fullCrewList) {
    return fullCrewList.filter(function (crew) {
      return crew.id === id;
    });
  }

  function getImmortalInfo(crew) {
    if (loadFrozen) {
      var symbol = findCrewById(crew.id, crewArchetypes.crew_avatars)[0].symbol,
        data = { 'symbol': symbol, 'access_token': access_token },
        options = {
          'method': 'post',
          'contentType': 'application/json',
          'payload': JSON.stringify(data)
        };
      var res = UrlFetchApp.fetch(apiDomain + 'stasis_vault/immortal_restore_info', options);
      return JSON.parse(res.getContentText()).crew;
    } else {
      var arch = findCrewById(crew.id, crewArchetypes.crew_avatars)[0];
      arch.archetype_id = arch.id;
      arch.rarity = 0;
      arch.level = 0;
      return arch;
    }
  }

  // haveState is 'Yes' for in roster, 'Vaulted' for immortals
  function appendCrewData(haveState) {
    // returns a function as a parameter for forEach
    return function (crew) {
      if (crew.in_buy_back_state) {
        // we don't care about dismissed crew
        return;
      }

      var id = crew.archetype_id;

      if (result.crew[crew.archetype_id].have) {
        // Duplicate crew!!!
        id = Math.floor(Math.random() * Math.floor(100000));
        result.crew[id] = JSON.parse(JSON.stringify(result.crew[crew.archetype_id]));
        result.crew[id].name = result.crew[id].name + " (DUPLICATE)";
      }

      result.crew[id].have = true;
      result.crew[id].haveState = haveState;
      result.crew[id].flavor = crew.flavor;
      result.crew[id].level = crew.level;
      result.crew[id].rarity = crew.rarity;
      result.crew[id].traits = crew.traits;
      result.crew[id].traits_hidden = crew.traits_hidden;
      result.crew[id].skills = crew.skills;
      result.crew[id].ship_battle = crew.ship_battle;
      result.crew[id].equipment = crew.equipment ? crew.equipment.length : 0;
      result.crew[id].favorite = crew.favorite;
      result.crew[id].airlocked = crew.in_buy_back_state;
    }
  }

  immortals = playerData.player.character.stored_immortals.map(getImmortalInfo);

  crewArchetypes = undefined;

  playerData.player.character.crew.forEach(appendCrewData('Yes'));
  immortals.forEach(appendCrewData('Vaulted'));

  playerData.player.character.cadet_schedule.missions.forEach(function (mission) {
    result.cadetMissions.push(mission.id);
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Crew roster");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('Crew roster');

  sheet.appendRow([' ', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  sheet.appendRow(["Name", "Have", "Rarity", "Max", "Level", "Equipment", 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', "Traits"]);

  var RARITYCOLORS = [
    { b: '', f: '', n: 'Basic' },
    { b: '#9b9b9b', f: '#080808', n: 'Common' },
    { b: '#50aa3c', f: '#963caa', n: 'Uncommon' },
    { b: '#5aaaff', f: '#ffaf5a', n: 'Rare' },
    { b: '#aa2deb', f: '#6eeb2d', n: 'Super Rare' },
    { b: '#fdd26a', f: '#6a95fd', n: 'Legendary' }
  ];

  var colIndex = 7;
  for (var skill in SKILLS) {
    var crew = SKILLS[skill];

    var range = sheet.getRange(1, colIndex, 1, 3);
    range.merge();
    range.setValue(SKILLS[skill]);
    range.setFontWeight("bold");
    range.setHorizontalAlignment("center");

    colIndex = colIndex + 3;
  }

  var crewArray = [];
  var backgroundColors = [];
  var fontColors = [];

  for (var crewId in result.crew) {
    var crew = result.crew[crewId];

    if (crew.have) {
      crewArray.push([
        crew.name,
        crew.haveState,
        crew.rarity,
        crew.max_rarity,
        crew.level,
        '' + crew.equipment + ' / 4',
        crew.skills.command_skill ? crew.skills.command_skill.core : 0,
        crew.skills.command_skill ? crew.skills.command_skill.range_min : 0,
        crew.skills.command_skill ? crew.skills.command_skill.range_max : 0,
        crew.skills.science_skill ? crew.skills.science_skill.core : 0,
        crew.skills.science_skill ? crew.skills.science_skill.range_min : 0,
        crew.skills.science_skill ? crew.skills.science_skill.range_max : 0,
        crew.skills.security_skill ? crew.skills.security_skill.core : 0,
        crew.skills.security_skill ? crew.skills.security_skill.range_min : 0,
        crew.skills.security_skill ? crew.skills.security_skill.range_max : 0,
        crew.skills.engineering_skill ? crew.skills.engineering_skill.core : 0,
        crew.skills.engineering_skill ? crew.skills.engineering_skill.range_min : 0,
        crew.skills.engineering_skill ? crew.skills.engineering_skill.range_max : 0,
        crew.skills.diplomacy_skill ? crew.skills.diplomacy_skill.core : 0,
        crew.skills.diplomacy_skill ? crew.skills.diplomacy_skill.range_min : 0,
        crew.skills.diplomacy_skill ? crew.skills.diplomacy_skill.range_max : 0,
        crew.skills.medicine_skill ? crew.skills.medicine_skill.core : 0,
        crew.skills.medicine_skill ? crew.skills.medicine_skill.range_min : 0,
        crew.skills.medicine_skill ? crew.skills.medicine_skill.range_max : 0,
        (crew.traits ? (crew.traits.join(', ') + ', ') : '') + (crew.traits_hidden ? (crew.traits_hidden.join(', ') + ', ') : '')
      ]);
    } else {
      crewArray.push([
        crew.name,
        'No',
        0,
        crew.max_rarity, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
        (crew.traits ? (crew.traits.join(', ') + ', ') : '') + (crew.traits_hidden ? (crew.traits_hidden.join(', ') + ', ') : '')
      ]);
    }

    backgroundColors.push(['white', 'white', RARITYCOLORS[crew.max_rarity].b, RARITYCOLORS[crew.max_rarity].b, 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white']);
    fontColors.push(['black', 'black', RARITYCOLORS[crew.max_rarity].f, RARITYCOLORS[crew.max_rarity].f, 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black', 'black']);
  }

  sheet.insertRows(3, crewArray.length);
  var crewRange = sheet.getRange(3, 1, crewArray.length, 25);
  crewRange.setValues(crewArray);
  crewRange.setBackgrounds(backgroundColors);
  crewRange.setFontColors(fontColors);

  // Freeze the first 2 rows
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  for (var i = 1; i < sheet.getLastColumn(); i++) {
    if (i > 6 && i < 25) {
      sheet.setColumnWidth(i, 38);
    } else {
      sheet.autoResizeColumn(i);
    }
  }

  colIndex = 7;
  for (var i = 0; i < 7; i++) {
    var range = sheet.getRange(1, colIndex, sheet.getLastRow());
    range.setBorder(null, true, null, null, false, false, null, null);
    colIndex = colIndex + 3;
  }

  sheet.showSheet();

  if (loadItems) {
    var sheet = ss.getSheetByName("Items");
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet('Items');
    sheet.appendRow(["Name", "Type", "Rarity", "Quantity", "Flavor"]);

    var itemArray = [];
    playerData.player.character.items.forEach(function (item) {
      itemArray.push([
        item.name,
        item.icon.file.replace("/items", "").split("/")[1],
        RARITYCOLORS[item.rarity].n,
        item.quantity,
        item.flavor ? item.flavor : '']);
    });

    sheet.insertRows(2, itemArray.length);
    var itemRange = sheet.getRange(2, 1, itemArray.length, 5);
    itemRange.setValues(itemArray);
    sheet.setFrozenRows(1);

    for (var i = 1; i < sheet.getLastColumn(); i++) {
      sheet.autoResizeColumn(i);
    }

    sheet.showSheet();
  }

  if (loadShips) {
    response = UrlFetchApp.fetch(apiDomain + 'ship_schematic' + apiQueryString);
    var shipData = JSON.parse(response.getContentText());

    var allShips = new Object();
    shipData.schematics.forEach(function (schematic) {
      allShips[schematic.ship.archetype_id] = schematic.ship;
      allShips[schematic.ship.archetype_id].level = 0;
    });

    Logger.log(allShips);

    playerData.player.character.ships.forEach(function (ship) {
      allShips[ship.archetype_id] = ship;
    });

    var sheet = ss.getSheetByName("Ships");
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet('Ships');
    sheet.appendRow(["Name", "Level", "Max Level", "Rarity", "Accuracy", "Antimatter", "Attack", "Attacks per second", "Crit Bonus", "Crit Chance", "Evasion", "Hull", "Shield Regen", "Shields", "Traits", "Flavor"]);

    var shipArray = [];
    for (var shipId in allShips) {
      var ship = allShips[shipId];
      shipArray.push([
        ship.name,
        ship.level,
        ship.max_level,
        ship.rarity,
        ship.accuracy,
        ship.antimatter,
        ship.attack,
        ship.attacks_per_second,
        ship.crit_bonus,
        ship.crit_chance,
        ship.evasion,
        ship.hull,
        ship.shield_regen,
        ship.shields,
        (ship.traits ? (ship.traits.join(', ') + ', ') : '') + (ship.traits_hidden ? (ship.traits_hidden.join(', ') + ', ') : ''),
        ship.flavor ? ship.flavor : '']);
    }

    sheet.insertRows(2, shipArray.length);
    var shipRange = sheet.getRange(2, 1, shipArray.length, 16);
    shipRange.setValues(shipArray);
    sheet.setFrozenRows(1);

    for (var i = 1; i < sheet.getLastColumn(); i++) {
      sheet.autoResizeColumn(i);
    }

    sheet.showSheet();
  }

  if (loadSnap) {
    var sheet = ss.getSheetByName("Stats");
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet('Stats');

    sheet.appendRow(["Name", playerData.player.character.display_name]);
    sheet.appendRow(["DBID", playerData.player.dbid]);
    sheet.appendRow(["Level", playerData.player.character.level]);
    sheet.appendRow(["VIP Level", playerData.player.vip_level]);
    sheet.appendRow(["XP", playerData.player.character.xp]);

    sheet.appendRow(["Honor", playerData.player.honor]);
    sheet.appendRow(["Dilithium", playerData.player.premium_purchasable]);
    sheet.appendRow(["Credits", playerData.player.money]);
    sheet.appendRow(["Merits", playerData.player.premium_earnable]);

    sheet.appendRow(["Crew Limit", playerData.player.character.crew_limit]);
    sheet.appendRow(["Shuttle Bays", playerData.player.character.shuttle_bays]);
    sheet.appendRow(["Max chronitons", playerData.player.character.replay_energy_max]);
    sheet.appendRow(["Chronitons in overflow", playerData.player.character.replay_energy_overflow]);
    sheet.appendRow(["Replicator uses today", "" + playerData.player.replicator_uses_today + " / " + playerData.player.replicator_limit]);

    sheet.appendRow(["Starbase buffs:"]);
    playerData.player.character.starbase_buffs.forEach(function (buff) {
      sheet.appendRow([buff.name, buff.stat]);
    });

    sheet.showSheet();
  }

  return result;
}

function loadCadetMissionData(cadetMissions) {
  var access_token = PropertiesService.getUserProperties().getProperty('accessToken');
  var missionIds = '';
  cadetMissions.forEach(function (missionId) {
    missionIds = missionIds + 'ids[]=' + missionId + '&';
  });

  var response = UrlFetchApp.fetch('https://stt.disruptorbeam.com/mission/info?' + missionIds + 'client_api=9&access_token=' + access_token);
  var missionData = JSON.parse(response.getContentText());

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cadet missions");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('Cadet missions');

  sheet.appendRow(["Mission", "Conflict", "Challenge", "Skill", "Difficulty", 'Crit', 'Traits', 'Trait bonus', 'Min Stars', 'Max Stars', 'Required Traits']);

  function getCrit(challenge) {
    if (!challenge.critical) {
      return 'None';
    }

    if (challenge.critical.claimed == true) {
      return 'Claimed (' + challenge.critical.threshold + ')';
    }

    return 'Unclaimed (' + challenge.critical.threshold + ')';
  }

  var questArray = [];
  var backgroundColors = [];

  missionData.character.accepted_missions.forEach(function (mission) {
    mission.quests.forEach(function (quest) {
      if (quest.quest_type == 'ConflictQuest') {
        response = UrlFetchApp.fetch('https://stt.disruptorbeam.com/quest/conflict_info?id=' + quest.id + '&client_api=9&access_token=' + access_token);
        var questData = JSON.parse(response.getContentText());

        if (questData.challenges.length == 0) {
          questArray.push([
            mission.episode_title,
            questData.name,
            'NO CHALLENGE YET',
            '', '', '', '', '',
            questData.crew_requirement.min_stars,
            questData.crew_requirement.max_stars,
            questData.crew_requirement.traits ? questData.crew_requirement.traits.join(', ') : ''
          ]);

          backgroundColors.push(['white', 'white', 'red', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white']);
        } else {
          questData.challenges.forEach(function (challenge) {
            var traits = [];
            var bonus = 0;
            challenge.trait_bonuses.forEach(function (traitBonus) {
              bonus = traitBonus.bonuses[2];
              traits.push(traitBonus.trait);
            });

            questArray.push([
              mission.episode_title,
              questData.name,
              challenge.name,
              SKILLS[challenge.skill],
              challenge.difficulty_by_mastery[2],
              getCrit(challenge),
              traits.join(', '),
              bonus,
              questData.crew_requirement.min_stars,
              questData.crew_requirement.max_stars,
              questData.crew_requirement.traits ? questData.crew_requirement.traits.join(', ') : ''
            ]);

            if (challenge.critical && challenge.critical.claimed == false) {
              backgroundColors.push(['white', 'white', 'white', 'white', 'white', 'red', 'white', 'white', 'white', 'white', 'white']);
            } else {
              backgroundColors.push(['white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white']);
            }
          });
        }
      }
    });
  });

  sheet.insertRows(3, questArray.length);
  var questRange = sheet.getRange(2, 1, questArray.length, 11);
  questRange.setValues(questArray);
  questRange.setBackgrounds(backgroundColors);

  // Freeze the first row
  sheet.setFrozenRows(1);

  for (var i = 1; i < sheet.getLastColumn(); i++) {
    sheet.autoResizeColumn(i);
  }

  sheet.showSheet();
}
