function getFastestLap(targetLocation, targetClass, cacheBuster) {
  //targetLocation = targetLocation || "Mt. Zion";
  //targetClass = targetClass || "LMP";
  cacheBuster = cacheBuster || new Date();
  var mtZion = {
    gt: [], 
    lmp: [], 
    club: []
  };
  var skybase = {
    gt: [], 
    lmp: [], 
    club: []
  };
  var fastestLap = null,
      location = null,
      class1 = null,
      class2 = null,
      sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var h = 0; h < sheets.length ; h++ ) {
    var sheet = sheets[h],
        checkSheet = sheet.getRange("A101").getValue();
    if (checkSheet !== "Points Race") {
      continue;
    }
    location = sheet.getRange("B3").getValue();
    class1 = sheet.getRange("A5").getValue();
    class2 = sheet.getRange("E5").getValue();
    var best1 = {
      name: sheet.getRange("A26").getValue(),
      time: sheet.getRange("B26").getValue()
    },
        best2 = {
          name: sheet.getRange("E26").getValue(),
          time: sheet.getRange("F26").getValue()
        }
    if (location === "Mt. Zion") {
      if (best1.name.length) {
        if (class1 === "GT") {
          mtZion.gt.push(best1);
        } else if (class1 === "LMP") {
          mtZion.lmp.push(best1);
        } else if (class1 === "Club Cars") {
          mtZion.club.push(best1);
        }
      }
      if (best2.name.length) {
        if (class2 === "GT") {
          mtZion.gt.push(best2);
        } else if (class2 === "LMP") {
          mtZion.lmp.push(best2);
        } else if (class2 === "Club Cars") {
          mtZion.club.push(best2);
        }
      }
    } else if (location === "Skybase 8011") {
      if (best1.name.length) {
        if (class1 === "GT") {
          skybase.gt.push(best1);
        } else if (class1 === "LMP") {
          skybase.lmp.push(best1);
        } else if (class1 === "Club Cars") {
          skybase.club.push(best1);
        }
      }
      if (best2.name.length) {
        if (class2 === "GT") {
          skybase.gt.push(best2);
        } else if (class2 === "LMP") {
          skybase.lmp.push(best2);
        } else if (class2 === "Club Cars") {
          skybase.club.push(best2);
        }
      }
    }
  }
  mtZion.gt = mtZion.gt.sort(compareLapTimes);
  mtZion.lmp = mtZion.lmp.sort(compareLapTimes);
  mtZion.club = mtZion.club.sort(compareLapTimes);
  skybase.gt = skybase.gt.sort(compareLapTimes);
  skybase.lmp = skybase.lmp.sort(compareLapTimes);
  skybase.club = skybase.club.sort(compareLapTimes);
  if (targetLocation === "Mt. Zion") {
    if (targetClass === "GT" && mtZion.gt.length) {
      fastestLap = mtZion.gt[0].name + ": " + mtZion.gt[0].time;
    } else if (targetClass === "LMP" && mtZion.lmp.length) {
      fastestLap = mtZion.lmp[0].name + ": " + mtZion.lmp[0].time;
    } else if (targetClass === "Club Cars" && mtZion.club.length) {
      fastestLap = mtZion.club[0].name + ": " + mtZion.club[0].time;
    }
  } else if (targetLocation === "Skybase 8011") {
    if (targetClass === "GT" && skybase.gt.length) {
      fastestLap = skybase.gt[0].name + ": " + skybase.gt[0].time;
    } else if (targetClass === "LMP" && skybase.lmp.length) {
      fastestLap = skybase.lmp[0].name + ": " + skybase.lmp[0].time;
    } else if (targetClass === "Club Cars" && skybase.club.length) {
      fastestLap = skybase.club[0].name + ": " + skybase.club[0].time;
    }
  }
  return fastestLap;
}

function compareLapTimes(a,b) {
  if (a.time < b.time)
    return -1;
  if (a.time > b.time)
    return 1;
  return 0;
}

function getTotalLaps(targetRider, targetClass, cacheBuster) {
  Logger.clear();
  Logger.log("exec: getTotalLaps");
  cacheBuster = cacheBuster || new Date();
  targetRider = targetRider || "Quinn";
  targetClass= targetClass || "LMP";
  var totalLaps = 0;
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  Logger.log("sheets.length: " + sheets.length);
  for (var h = 0; h < sheets.length ; h++ ) {
    var sheet = sheets[h];
    var checkSheet = sheet.getRange("A101").getValue();
    Logger.log("checkSheet: " + checkSheet);
    if (checkSheet === "Points Race") {
      Logger.log("Score sheet: " + sheet.getName()); 
    } else {
      Logger.log("NOT score sheet: " + sheet.getName()); 
      continue;
    }
    var range1 = sheet.getRange(5, 1, 20, 3);
    var range2 = sheet.getRange(5, 5, 20, 3);
    totalLaps = totalLaps + processScores2(range1, targetRider, targetClass, cacheBuster);
    totalLaps = totalLaps + processScores2(range2, targetRider, targetClass, cacheBuster);
  }
  Logger.log("test: " + totalLaps);
  return totalLaps;
}

function processScores2(range, targetRider, targetClass, cacheBuster) {
  var mode = null;
  var grabLaps = false;
  var rangeLaps = 0;
  var vals = range.getValues();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      //var currentValue = range.getCell(i,j).getValue();
      var currentValue = vals[i-1][j-1];
      //Logger.log("======");
      //Logger.log("currentValue: " + currentValue);
      //Logger.log("currentValue2: " + currentValue2);
      //Logger.log("======");
      //if (currentValue.length) {
        //Logger.log(currentValue);
      //}
      if (grabLaps === true) {
        Logger.log("Adding laps: " + currentValue);
        rangeLaps = rangeLaps + currentValue; 
        grabLaps = false;
      }
      if (currentValue === "GT") {
        mode = "GT";
      } else if (currentValue === "LMP") {
        mode = "LMP"
      } else if (currentValue === "Club Cars") {
        mode = "Club Cars"
      }
      if (mode === targetClass && currentValue === targetRider) {
        grabLaps = true; 
      }
    }
  }
  return rangeLaps;
}

function processScores(range, targetRider, targetClass, cacheBuster) {
  var mode = null;
  var grabLaps = false;
  var rangeLaps = 0;
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var currentValue = range.getCell(i,j).getValue();
      if (currentValue.length) {
        //Logger.log(currentValue);
      }
      if (grabLaps === true) {
        Logger.log("Adding laps: " + currentValue);
        rangeLaps = rangeLaps + currentValue; 
        grabLaps = false;
      }
      if (currentValue === "GT") {
        mode = "GT";
      } else if (currentValue === "LMP") {
        mode = "LMP"
      } else if (currentValue === "Club Cars") {
        mode = "Club Cars"
      }
      if (mode === targetClass && currentValue === targetRider) {
        grabLaps = true; 
      }
    }
  }
  return rangeLaps;
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Force Refresh",
    functionName : "refreshLastUpdate"
  }];
  sheet.addMenu("Tiny Riders", entries);
};

function refreshLastUpdate() {
  SpreadsheetApp.getActiveSpreadsheet().getRange('A102').setValue(new Date().toTimeString());
}