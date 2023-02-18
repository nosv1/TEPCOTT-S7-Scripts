import 'google-apps-script';


/******************************************************************************/
// Constants
/******************************************************************************/


// position backgrounds
const POINTS_PROMOTED_BG = "#d9ead3";
const POINTS_DEMOTED_BG = "#f4cccc";
const POINTS_STAYED_BG = "#d5e3f5";
const ROUND_TAB_PROMOTED_BG = "#d9ead3"; // the color of the background that corresponds to a promotion granting position
const ROUND_TAB_DEMOTED_BG = "#f4cccc";  // the color of the background that corresponds to a demotion granting position
const ROUND_TAB_STAYED_BG = "#abc7eb";   // the color of the background that corresponds to a stayed in division position
const ROUND_TAB_DROPPED_BG = "#ff0000";  // the color of the background that corresponds to a dropped out driver

// round tab named ranges that return a1 notation ranges
const ROUND_TAB_DIVISION_ROW_OFFSET_NAMED_RANGE = "my_sheet_round_tab_division_offset";
const ROUND_TAB_EXTRA_INFORMATION_NAMED_RANGE = "my_sheet_finishing_order_extra_information_range";
const ROUND_TAB_FINISHING_ORDER_DRIVERS_NAMED_RANGE = "my_sheet_finishing_order_drivers_range";
const ROUND_TAB_FINISHING_ORDER_RESERVES_NAMED_RANGE = "my_sheet_finishing_order_reserves_range";
const ROUND_TAB_STARTING_ORDER_DRIVERS_NAMED_RANGE = "my_sheet_starting_order_drivers_range";
const ROUND_TAB_STARTING_ORDER_DRIVER_DIVISIONS_NAMED_RANGE = "my_sheet_starting_order_divisions_range"
const ROUND_TAB_STARTING_ORDER_DRIVERS_POINTS_NAMED_RANGE = "my_sheet_starting_order_drivers_points_range";
const ROUND_TAB_STARTING_ORDER_RESERVES_NAMED_RANGE = "my_sheet_starting_order_reserves_range";
const ROUND_TAB_STARTING_ORDER_RESERVE_DIVISIONS_NAMED_RANGE = "my_sheet_starting_order_reserve_divs_range";
const ROUND_TAB_TIEBREAKER_INDICATOR = "Spin Wheel!";
const ROUND_TAB_RESERVE_NEEDED_STRING = "RESERVE NEEDED";
const ROUND_TAB_WAITING_LIST_DRIVERS_NAMED_RANGE = "my_sheet_waiting_list_drivers_range";
const ROUND_TAB_WAITING_LIST_ACCEPTABLE_DIVISIONS_NAMED_RANGE = "my_sheet_waiting_list_acceptable_divisions_range";

const ROUND_TAB_DROPPING_OUT_EXTRA_INFORMATION = "dropping out";
ROUND_TAB_SHEET_NAME_PATTERN = /r\d+/;
const ROUND_TAB_PREFIX = 'r'
const WAITING_LIST_DIVISION_STRING = "WL";

// points tab named ranges
const POINTS_SHEET_NAME = "points";
const POINTS_DIV_1_POINTS = "points_div_1_points";
const POINTS_DIV_2_POINTS = "points_div_2_points";
const POINTS_DIV_3_POINTS = "points_div_3_points";
const POINTS_DIV_4_POINTS = "points_div_4_points";
const POINTS_DIV_5_POINTS = "points_div_5_points";
const POINTS_DIV_6_POINTS = "points_div_6_points";

// leaderboard named ranges
const LEADERBOARD_SHEET_NAME = "leaderboard";
const LEADERBOARD_DRIVERS = "leaderboard_drivers";
const LEADERBOARD_POINTS = "leaderboard_total_points";

// starting order by movement
const MOVEMENTS = {
  PROMOTED: "PROMOTED",
  DEMOTED: "DEMOTED",
  STAYED: "STAYED",
  FROM_WAITING_LIST: "FROM_WAITING_LIST",
  DROPPED: "DROPPED"
}
const STARTING_ORDER_BY_MOVEMENT = [
  MOVEMENTS.STAYED, 
  MOVEMENTS.DEMOTED, 
  MOVEMENTS.PROMOTED, 
  MOVEMENTS.FROM_WAITING_LIST
]


/******************************************************************************/
// Classes
/******************************************************************************/


/**
 * @class Driver
 * @description This class represents a driver
 * @param name {string} The name of the driver
 * @param division {int} The division the driver is in
 * @param finishing_position {int} The finishing position of the driver
 * @param reserve {Driver} The name of the reserve driver
 * @param extra_information {string} Any extra information about the driver
 */
class Driver {
  constructor(name, division = null, finishing_position = null, reserve = null, extra_information = null, ) {
    this.name = name;
    this.division = division;
    this.finishing_position = finishing_position;
    this.reserve = reserve;
    this.extra_information = extra_information;

    this.promoted = false;
    this.demoted = false;
    this.stayed = false;
    this.from_waiting_list = false;

    this.total_points = null;
    this.tied = false;
  }

  /**
   * @returns {boolean}
   * @description This function returns true if the driver is dropping out, false otherwise
   */
  get has_dropped() {
    return this.extra_information == ROUND_TAB_DROPPING_OUT_EXTRA_INFORMATION
  }

  /**
   * @returns {string} The movement of the driver
   * @description This function returns the movement of the driver
   */
  get movement() {
    if (this.has_dropped) {
      return MOVEMENTS.DROPPED;
    } else if (this.promoted) {
      return MOVEMENTS.PROMOTED;
    } else if (this.demoted) {
      return MOVEMENTS.DEMOTED;
    } else if (this.stayed) {
      return MOVEMENTS.STAYED;
    } else if (this.from_waiting_list) {
      return MOVEMENTS.FROM_WAITING_LIST;
    }
  }

  /**
   * @returns {string} The background color of the driver
   * @description This function returns the background color associated with if the driver is promoted, demoted, stayed, or dropped
   */
  get background() {
    if (this.has_dropped) {
      return ROUND_TAB_DROPPED_BG;
    } else if (this.promoted) {
      return ROUND_TAB_PROMOTED_BG;
    } else if (this.demoted) {
      return ROUND_TAB_DEMOTED_BG;
    } else if (this.stayed) {
      return ROUND_TAB_STAYED_BG;
    } else if (this.from_waiting_list) {
      return ROUND_TAB_STAYED_BG;
    } else {
      return ROUND_TAB_STAYED_BG;
    }
  }

  /**
   * @param leaderboard_drivers {string[][]}
   * @param leaderboard_points {int[][]}
   */
  setPoints(leaderboard_drivers, leaderboard_points) {
    let driver_index = leaderboard_drivers.findIndex(driver => driver[0] == this.name);
    this.total_points = leaderboard_points[driver_index][0];
  }
}


/**
 * @class Division
 * @description This class represents a division
 * @param division_number {int} The number of the division
 * @param drivers {Driver[]} The drivers in the division
 * @param waiting_list {Driver[]} The drivers on the waiting list
 */
class Division {
  constructor(division_number, drivers = [], waiting_list = []) {
    /**
     * @type {int}
     */
    this.division_number = division_number;
    /**
     * @type {Driver[]}
     */
    this.drivers = drivers;
    /**
     * @type {Driver[]}
     */
    this.waiting_list = waiting_list;
  }
}


/******************************************************************************/
// Utility Functions
/******************************************************************************/


/**
 * @param division_number {int} The number of the division
 * @param drivers {Driver[]} The drivers in the division
 */
function sortStartingOrder(division_number, drivers) {
  // div 1 is sorted reverse points order (ascending) of drivers that were not from waiting list
  // all other divs are sorted based on movement then by points
  
  let reserves_faster_than_div = [];
  let i = drivers.length;
  while (true) {
    i -= 1;
    if (i < 0) {
      break;
    }

    let driver = drivers[i];
    if (driver.reserve != null && driver.reserve.division < division_number) {
      reserves_faster_than_div.push(driver);
      drivers.splice(i, 1);
    }
  }

  let drivers_from_waiting_list = [];
  i = drivers.length;
  while (true) {
    i -= 1;
    if (i < 0) {
      break;
    }

    let driver = drivers[i];
    if (driver.from_waiting_list) {
      driver.tied = true;
      drivers_from_waiting_list.push(driver);
      drivers.splice(i, 1);
    }
  }
  
  /**
   * @param a {Driver}
   * @param b {Driver}
   */
  function handleTies(a, b) {
    if (a.total_points == b.total_points) {
      a.tied = true;
      b.tied = true;
    }
  }

  if (division_number == 1) {
    drivers.sort(function(a, b) {
      handleTies(a, b);
      return a.total_points - b.total_points;
    });
  } else {
    drivers.sort(function(a, b) {
      if (a.movement != b.movement) {
        return STARTING_ORDER_BY_MOVEMENT.indexOf(a.movement) - STARTING_ORDER_BY_MOVEMENT.indexOf(b.movement);
      }

      handleTies(a, b);
      return a.total_points - b.total_points;
    });
  }

  for (let i = reserves_faster_than_div.length - 1; i >= 0; i--) {
    let reserve = reserves_faster_than_div[i];
    drivers.push(reserve);
  }

  for (let i = drivers_from_waiting_list.length - 1; i >= 0; i--) {
    let driver = drivers_from_waiting_list[i];
    drivers.push(driver);
  }
  return drivers;
}


/**
 * @param alert_text {string} The text to display in the alert
 * @returns {boolean} True if the user clicked yes, false otherwise
 * @description This function displays an alert to the user and returns true if the user clicked yes, false otherwise
 */
function confirm_continue(alert_text) {
  let result = SpreadsheetApp.getUi().alert(
    alert_text,
    SpreadsheetApp.getUi().ButtonSet.YES_NO)
  return result == SpreadsheetApp.getUi().Button.YES
}


/**
 * @param values {Array} An array of columns from a range.getValues() call
 * @param division_row_offset {number} The number of rows to skip before starting a new division
 * @returns {Array} The values split by division offset
 * @description This function splits a range into divisions based on the division_row_offset
 */
function splitArrayOfColumnsIntoDivisions(values, division_row_offset) {
  let divisions = [];
  let division = [];
  for (let i = 0; i < values[0].length; i++) {
    if (i % division_row_offset == 0 && i != 0) {
      divisions.push(division);
      division = [];
    }
    let row = []
    for (let j = 0; j < values.length; j++) {
      row.push(values[j][i])
    }
    division.push(row);
  }
  if (division.length > 0) {
    divisions.push(division);
  }
  return divisions;
}

/**
 * @param driver {Driver} The driver to update
 * @param previous_division {str} The previous division of the driver
 * @returns {driver} The updated driver
 */
function updateDriverMovementFromDivision(driver, previous_division) {
  if (previous_division == WAITING_LIST_DIVISION_STRING) {
    driver.from_waiting_list = true;
    return driver;
  }

  previous_division = parseInt(previous_division);
  if (previous_division < driver.division) {
    driver.demoted = true;
  } else if (previous_division > driver.division) {
    driver.promoted = true;
  } else {
    driver.stayed = true;
  }
  return driver;
}


/******************************************************************************/
// Update Divisions
/******************************************************************************/



/**
 * @param all_divisions {Division[]}
 * @returns {Object{Division[], Division[]}} 
 */
function handleDropouts(all_divisions) {
  
  /**
   * @param divisions {Division[]}
   * @param start {int}
   * @param end {int}
   */
  function labelHighestFastestDemotionAsStayed(divisions, start, end) {

    for (let i = start; i < end; i++) {
      let division = divisions[i];

      if (division.waiting_list.length > 0) {
        let waiting_list_driver = division.waiting_list.shift();
        waiting_list_driver.from_waiting_list = true;
        division.drivers.push(waiting_list_driver);

        // remove waiting list driver from any other waiting lists he may be in
        for (let k = 0; k < divisions.length; k++) {
          let other_division = divisions[k];
          for (let l = 0; l < other_division.waiting_list.length; l++) {
            let other_waiting_list_driver = other_division.waiting_list[l];
            if (other_waiting_list_driver.name == waiting_list_driver.name) {
              other_division.waiting_list.splice(l, 1);
              break;
            }
          }
        }
        break;
      }

      for (let j = 0; j < division.drivers.length; j++) {
        let driver = division.drivers[j];
        if (driver.demoted) {
          driver.stayed = true
          driver.demoted = false
          break;
        }
      }
    }
    return divisions;
  }

  for (let i = 0; i < all_divisions.length; i++) {
    let division = all_divisions[i];

    for (let j = 0; j < division.drivers.length; j++) {
      let driver = division.drivers[j];

      if (driver.has_dropped) {
        if (driver.promoted) {
          all_divisions = labelHighestFastestDemotionAsStayed(all_divisions, i-1, all_divisions.length - 1);
        } else if (driver.demoted) {
          all_divisions = labelHighestFastestDemotionAsStayed(all_divisions, i+1, all_divisions.length - 1);
        } else if (driver.stayed) {
          all_divisions = labelHighestFastestDemotionAsStayed(all_divisions, i, all_divisions.length - 1);
        }
      }
    }
  }
  return all_divisions;
}


function updateDivisions() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // get the active spreadsheet
  let sheet = SpreadsheetApp.getActiveSheet(); // get the active sheet
  let sheet_name = sheet.getName(); // get the name of the active sheet

  if (!confirm_continue(
      `Create a new division based on \`${sheet_name}\`? \
          \n\nHint: This should be the current round tab's name...`)) {
    console.log("cancelling division update")
    return;
  }
  
  let round_number = parseInt(sheet_name.match(/\d+/)[0]);

  // get round tab ranges
  let round_tab_named_ranges = [
    ROUND_TAB_FINISHING_ORDER_DRIVERS_NAMED_RANGE,
    ROUND_TAB_FINISHING_ORDER_RESERVES_NAMED_RANGE,
    ROUND_TAB_EXTRA_INFORMATION_NAMED_RANGE,
    ROUND_TAB_WAITING_LIST_DRIVERS_NAMED_RANGE,
    ROUND_TAB_WAITING_LIST_ACCEPTABLE_DIVISIONS_NAMED_RANGE
  ]
  let finishing_order_driver_index = round_tab_named_ranges.indexOf(ROUND_TAB_FINISHING_ORDER_DRIVERS_NAMED_RANGE);
  let finishing_order_reserve_index = round_tab_named_ranges.indexOf(ROUND_TAB_FINISHING_ORDER_RESERVES_NAMED_RANGE);
  let finishing_order_extra_information_index = round_tab_named_ranges.indexOf(ROUND_TAB_EXTRA_INFORMATION_NAMED_RANGE);
  let waiting_list_drivers_index = round_tab_named_ranges.indexOf(ROUND_TAB_WAITING_LIST_DRIVERS_NAMED_RANGE);
  let waiting_list_acceptable_divisions_index = round_tab_named_ranges.indexOf(ROUND_TAB_WAITING_LIST_ACCEPTABLE_DIVISIONS_NAMED_RANGE);

  let round_tab_a1_ranges = round_tab_named_ranges.map(r => spreadsheet.getRangeByName(r).getValue());
  //// above is gross... it gets the ranges one at a time, but to use .getRangeList you have to know the source sheet, and I didn't want to define what sheet the named ranges are located in
  let round_tab_ranges = sheet.getRangeList(round_tab_a1_ranges).getRanges();

  let finishing_order_range = round_tab_ranges[finishing_order_driver_index];
  let reserve_range = round_tab_ranges[finishing_order_reserve_index];
  let extra_information_range = round_tab_ranges[finishing_order_extra_information_index];
  let waiting_list_drivers_range = round_tab_ranges[waiting_list_drivers_index];
  let waiting_list_acceptable_divisions_range = round_tab_ranges[waiting_list_acceptable_divisions_index];

  let round_tab_values = [
    finishing_order_range.getValues(),
    reserve_range.getValues(),
    extra_information_range.getValues(),
    waiting_list_drivers_range.getValues(),
    waiting_list_acceptable_divisions_range.getValues().map(v => v[0].split(',').map(v => parseInt(v))) // div range is a column of `div, div` strings, so we split them and convert to ints
  ]

  // get points ranges
  let points_named_ranges = [
    POINTS_DIV_1_POINTS,
    POINTS_DIV_2_POINTS,
    POINTS_DIV_3_POINTS,
    POINTS_DIV_4_POINTS,
    POINTS_DIV_5_POINTS,
    POINTS_DIV_6_POINTS,
  ]
  let points_sheet = spreadsheet.getSheetByName(POINTS_SHEET_NAME);
  let points_ranges = points_sheet.getRangeList(points_named_ranges).getRanges();
  let points_backgrounds = points_ranges.map(r => r.getBackgrounds());

  // get leaderboard ranges
  let leaderboard_named_ranges = [
    LEADERBOARD_DRIVERS,
    LEADERBOARD_POINTS,
  ]
  let leaderboard_drivers_index = leaderboard_named_ranges.indexOf(LEADERBOARD_DRIVERS);
  let leaderboard_points_index = leaderboard_named_ranges.indexOf(LEADERBOARD_POINTS);
  let leaderboard_sheet = spreadsheet.getSheetByName(LEADERBOARD_SHEET_NAME);
  let leaderboard_ranges = leaderboard_sheet.getRangeList(leaderboard_named_ranges).getRanges();
  let leaderboard_values = [
    leaderboard_ranges[leaderboard_drivers_index].getValues(),
    leaderboard_ranges[leaderboard_points_index].getValues(),
  ]

  let next_round_tab_named_ranges = [
    ROUND_TAB_STARTING_ORDER_DRIVERS_NAMED_RANGE,
    ROUND_TAB_STARTING_ORDER_RESERVES_NAMED_RANGE,
    ROUND_TAB_EXTRA_INFORMATION_NAMED_RANGE,
  ]
  let starting_order_driver_index = next_round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_DRIVERS_NAMED_RANGE);
  let starting_order_reserve_index = next_round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_RESERVES_NAMED_RANGE);
  let starting_order_extra_information_index = next_round_tab_named_ranges.indexOf(ROUND_TAB_EXTRA_INFORMATION_NAMED_RANGE);
  let next_round_tab_a1_ranges = next_round_tab_named_ranges.map(r => spreadsheet.getRangeByName(r).getValue())
  
  let next_round_sheet = spreadsheet.getSheetByName(`${ROUND_TAB_PREFIX}${round_number + 1}`);
  let next_round_tab_ranges = next_round_sheet.getRangeList(next_round_tab_a1_ranges).getRanges();
  let next_round_tab_values = [
    next_round_tab_ranges[starting_order_driver_index].getValues(),
    next_round_tab_ranges[starting_order_reserve_index].getValues(),
    next_round_tab_ranges[starting_order_extra_information_index].getValues(),
  ]
  let next_round_tab_backgrounds = [
    next_round_tab_ranges[starting_order_driver_index].getBackgrounds()
  ]


  // split into divisions
  let division_row_offset = sheet.getRange(ROUND_TAB_DIVISION_ROW_OFFSET_NAMED_RANGE).getValue();
  let _divisions = splitArrayOfColumnsIntoDivisions([
    round_tab_values[finishing_order_driver_index],
    round_tab_values[finishing_order_reserve_index],
    round_tab_values[finishing_order_extra_information_index]
  ], division_row_offset);

  // convert to division and driver objects
  /**
   * @type {Division[]}
   */
  let divisions = [];
  for (let i = 0; i < _divisions.length; i++) {
    let drivers = _divisions[i];
    let _division = new Division(
      division_number = i + 1,
    )
    let reserves_promoted = 0;  // count of would be promotions if not for reserves

    for (let j = 0; j < _divisions[i].length; j++) {
      let driver_name = drivers[j][finishing_order_driver_index][0]
      let reserve_name = drivers[j][finishing_order_reserve_index][0]
      if (driver_name == '') {
        continue;
      }

      let driver = new Driver(
        name = driver_name,
        division = _division.division_number,
        finishing_position = j + 1,
        reserve = reserve_name != '' ? new Driver(reserve_name) : null,
        extra_information = drivers[j][finishing_order_extra_information_index][0]
      )

      let reserve_promotion = driver.reserve != null && points_backgrounds[i][j - reserves_promoted][0] == ROUND_TAB_PROMOTED_BG
      if (reserve_promotion) {
        reserves_promoted++;
      }

      driver.promoted = !reserve_promotion && points_backgrounds[i][j - reserves_promoted][0] == POINTS_PROMOTED_BG
      driver.demoted = points_backgrounds[i][j][0] == POINTS_DEMOTED_BG
      driver.stayed = reserve_promotion || points_backgrounds[i][j][0] == POINTS_STAYED_BG
      _division.drivers.push(driver);
    }
    divisions.push(_division);
  }

  // populate div waiting lists
  for (let i = 0; i < round_tab_values[waiting_list_drivers_index].length; i++) {
    let driver_name = round_tab_values[waiting_list_drivers_index][i][0];
    if (driver_name == "") {
      continue;
    }
    let driver = new Driver(name=driver_name)
    let acceptable_divisions = round_tab_values[waiting_list_acceptable_divisions_index][i];
    for (let j = 0; j < acceptable_divisions.length; j++) {
      let acceptable_division = acceptable_divisions[j];
      divisions[acceptable_division - 1].waiting_list.push(driver);
    }
  }

  // handle dropouts
  divisions = handleDropouts(divisions);
  let updated_divisions = [new Division(1)];
  
  // set finishing order background colors and update division for next round
  let finishing_order_backgrounds = finishing_order_range.getBackgrounds();
  for (let i = 0; i < divisions.length; i++) {
    let division = divisions[i];
    updated_divisions.push(new Division(division.division_number + 1))

    let j = division.drivers.length - 1;
    while (j >= 0) {
      let driver = division.drivers[j];
      finishing_order_backgrounds[i * division_row_offset + j][0] = driver.background;

      if (driver.has_dropped) {
        division.drivers.splice(j, 1);
        j--;
        continue;
      }

      driver.setPoints(
        leaderboard_values[leaderboard_drivers_index], 
        leaderboard_values[leaderboard_points_index]
      );

      driver.reserve = null;
      
      if (driver.promoted) {
        updated_divisions[division.division_number - 2].drivers.push(driver);
      } else if (driver.demoted) {
        updated_divisions[division.division_number].drivers.push(driver);
      } else if (driver.stayed) {
        updated_divisions[division.division_number - 1].drivers.push(driver);
      } else if (driver.from_waiting_list) {
        updated_divisions[division.division_number - 1].drivers.push(driver);
      }

      j--;
    }
  }
  updated_divisions.pop();
  finishing_order_range.setBackgrounds(finishing_order_backgrounds);

  for (let i = 0; i < updated_divisions.length; i++) {
    let division_number = i + 1;
    updated_divisions[i].drivers = sortStartingOrder(division_number, updated_divisions[i].drivers);
    while (updated_divisions[i].drivers.length < division_row_offset) {
      updated_divisions[i].drivers.push(new Driver(""));
    }

    for (let j = 0; j < updated_divisions[i].drivers.length - 1; j++) {
      let driver = updated_divisions[i].drivers[j];
      next_round_tab_values[starting_order_driver_index][i * division_row_offset + j][0] = driver.name;
      next_round_tab_values[starting_order_reserve_index][i * division_row_offset + j][0] = driver.reserve ? driver.reserve.name : "";
      next_round_tab_backgrounds[starting_order_driver_index][i * division_row_offset + j][0] = driver.background;
      if (driver.tied) {
        next_round_tab_values[starting_order_extra_information_index][i * division_row_offset + j][0] = ROUND_TAB_TIEBREAKER_INDICATOR
      } else {
        next_round_tab_values[starting_order_extra_information_index][i * division_row_offset + j][0] = ""
      }
    }
  }
  
  next_round_tab_ranges[starting_order_driver_index].setValues(next_round_tab_values[starting_order_driver_index]);
  next_round_tab_ranges[starting_order_reserve_index].setValues(next_round_tab_values[starting_order_reserve_index]);
  next_round_tab_ranges[starting_order_driver_index].setBackgrounds(next_round_tab_backgrounds[starting_order_driver_index]);
  next_round_tab_ranges[starting_order_extra_information_index].setValues(next_round_tab_values[starting_order_extra_information_index]);
}


/******************************************************************************/
// Update Starting Orders
/******************************************************************************/

function updateStartingOrders() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // get the active spreadsheet
  let sheet = SpreadsheetApp.getActiveSheet(); // get the active sheet
  let sheet_name = sheet.getName(); // get the name of the active sheet

  if (!confirm_continue(
      `Create a new division based on \`${sheet_name}\`? \
          \n\nHint: This should be the current round tab's name...`)) {
    console.log("cancelling division update")
    return;
  }
  
  let round_number = parseInt(sheet_name.match(/\d+/)[0]);

  let round_tab_named_ranges = [
    ROUND_TAB_STARTING_ORDER_DRIVER_DIVISIONS_NAMED_RANGE,
    ROUND_TAB_STARTING_ORDER_DRIVERS_NAMED_RANGE,
    ROUND_TAB_STARTING_ORDER_DRIVERS_POINTS_NAMED_RANGE,
    ROUND_TAB_STARTING_ORDER_RESERVE_DIVISIONS_NAMED_RANGE,
    ROUND_TAB_STARTING_ORDER_RESERVES_NAMED_RANGE
  ]
  let starting_order_driver_division_index = round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_DRIVER_DIVISIONS_NAMED_RANGE);
  let starting_order_driver_index = round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_DRIVERS_NAMED_RANGE);
  let starting_order_driver_points_index = round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_DRIVERS_POINTS_NAMED_RANGE);
  let starting_order_reserve_division_index = round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_RESERVE_DIVISIONS_NAMED_RANGE);
  let starting_order_reserve_index = round_tab_named_ranges.indexOf(ROUND_TAB_STARTING_ORDER_RESERVES_NAMED_RANGE);

  let round_tab_a1_ranges = round_tab_named_ranges.map(r => spreadsheet.getRangeByName(r).getValue());
  let round_tab_ranges = sheet.getRangeList(round_tab_a1_ranges).getRanges();

  let starting_order_driver_divisions_range = round_tab_ranges[starting_order_driver_division_index];
  let starting_order_drivers_range = round_tab_ranges[starting_order_driver_index];
  let starting_order_driver_points_range = round_tab_ranges[starting_order_driver_points_index];
  let starting_order_reserve_divisions_range = round_tab_ranges[starting_order_reserve_division_index];
  let starting_order_reserves_range = round_tab_ranges[starting_order_reserve_index];

  let round_tab_values = [
    starting_order_driver_divisions_range.getValues(),
    starting_order_drivers_range.getValues(),
    starting_order_driver_points_range.getValues(),
    starting_order_reserve_divisions_range.getValues(),
    starting_order_reserves_range.getValues()
  ]

  let starting_order_drivers_backgrounds = starting_order_drivers_range.getBackgrounds();

  // split into divisions
  let division_row_offset = sheet.getRange(ROUND_TAB_DIVISION_ROW_OFFSET_NAMED_RANGE).getValue();
  let _divisions = splitArrayOfColumnsIntoDivisions([
    round_tab_values[starting_order_driver_division_index],
    round_tab_values[starting_order_driver_index],
    round_tab_values[starting_order_driver_points_index],
    round_tab_values[starting_order_reserve_division_index],
    round_tab_values[starting_order_reserve_index]
  ], division_row_offset);


  for (let i = 0; i < _divisions.length; i++) {
    let drivers = _divisions[i]
    let _division = new Division(
      division_number = i + 1,
    );

    for (let j = 0; j < drivers.length; j++) {
      let driver_name = drivers[j][starting_order_driver_index][0];
      let reserve_name = drivers[j][starting_order_reserve_index][0];
      let reserve_division = drivers[j][starting_order_reserve_division_index][0];
      if (driver_name == "") {
        continue;
      }
      
      let reserve = null;
      if (reserve_name == ROUND_TAB_RESERVE_NEEDED_STRING) {
        reserve = new Driver(name=reserve_name, division=_division.division_number);
      } else if (reserve_name != "") {
        let reserve_division_number = parseInt(String(reserve_division).match(/\d+/)[0]);
        reserve = new Driver(
          name = reserve_name,
          division = reserve_division_number
        );
      }


      let _driver = drivers[j];
      let driver = new Driver(
        name = _driver[starting_order_driver_index][0],
        division = _division.division_number,
        points = _driver[starting_order_driver_points_index][0],
        reserve = reserve,
      );
      driver = updateDriverMovementFromDivision(driver, drivers[j][starting_order_driver_division_index][0]);
      _division.drivers.push(driver);
    }
    _division.drivers = sortStartingOrder(_division.division_number, _division.drivers);

    while (_division.drivers.length < division_row_offset) {
      _division.drivers.push(new Driver(name = ""));
    }
    
    // set the range values and backgrounds for staring order drivers and starting order reserves
    for (let j = 0; j < _division.drivers.length - 1; j++) {
      let driver = _division.drivers[j];
      let reserve = driver.reserve;
      if (reserve == null) {
        reserve = new Driver(name = "");
      }
      round_tab_values[starting_order_driver_index][i * division_row_offset + j][0] = driver.name;
      starting_order_drivers_backgrounds[i * division_row_offset + j][0] = driver.background;
      round_tab_values[starting_order_reserve_index][i * division_row_offset + j][0] = reserve.name;
    }
  }

  starting_order_drivers_range.setBackgrounds(starting_order_drivers_backgrounds);
  starting_order_drivers_range.setValues(round_tab_values[starting_order_driver_index]);
  starting_order_reserves_range.setValues(round_tab_values[starting_order_reserve_index]);
}