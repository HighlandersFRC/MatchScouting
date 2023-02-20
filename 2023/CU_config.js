var config_data = `
{
  "title": "Scouting Site 2023",
  "page_title": "Charged Up",
  "checkboxAs": "10",
  "prematch": [
    { "name": "Scouter Initials",
      "code": "s",
      "type": "scouter",
      "size": 5,
      "maxSize": 5,
      "required": "true"
    },
    { "name": "Event",
      "code": "e",
      "type": "event",
      "defaultValue": "2022carv",
      "required": "true",
      "disabled": "true"
    },
    { "name": "Match #",
      "code": "m",
      "type": "match",
      "min": 1,
      "max": 1000,
      "required": "true"
    },
    { "name": "Match Level",
      "code": "l",
      "type": "level",
      "choices": {
        "qm": "Quals<br>"
      },
      "defaultValue": "qm",
      "required": "true",
      "disabled": "true"
    },
    { "name": "Robot",
      "code": "r",
      "type": "robot",
      "choices": {
        "r1": "Red-1",
        "b1": "Blue-1<br>",
        "r2": "Red-2",
        "b2": "Blue-2<br>",
        "r3": "Red-3",
        "b3": "Blue-3"
      },
      "required":"true"
    },
    { "name": "Team #",
      "code": "t",
      "type": "team",
      "min": 1,
      "max": 99999
    },
    { "name": "Auto Start Position",
      "code": "as",
      "type": "radio",
      "filename": "2023/field_image_updated.png",
      "choices": {
        "1": "1<br>",
        "2": "2<br>",
        "3": "3"
      }
    }
  ],
  "auton": [
    { "name": "Auto Scoring",
      "code": "asg",
      "type": "clickable_image",
      "filename": "2023/grid_image.png",
      "dimensions": "9 3",
      "clickRestriction": "onePerBox",
      "toggleClick": "true",
      "showFlip": "false",
      "showUndo": "false",
      "shape": "circle 12 black red true"
    },
    { "name": "Exited Community",
      "code": "ec",
      "type": "bool"
    },
    { "name": "Docked",
      "code": "ad",
      "type":"radio",
      "choices": {
        "dock": "Docked (not Engaged)<br>",
        "eng": "Engaged<br>",
        "att": "Attempted but failed<br>",
        "x": "Not attempted"
      },
      "defaultValue": "x"
    },
    { "name": "Game Pieces Attempted",
      "code": "agpa",
      "type": "counter",
      "required": "true",
      "defaultValue": 0,
      "max": 5,
      "min": 0
    },
    { "name": "Holding, but not<br>placing a piece",
      "code": "gph",
      "type": "bool"
    }
  ],
  "teleop": [
    { "name": "Cycle Timer",
      "code": "tct",
      "type": "cycle"
    },
    { "name": "Grid Scoring",
      "code": "tsg",
      "type": "clickable_image",
      "filename": "2023/grid_image.png",
      "dimensions": "9 3",
      "clickRestriction": "onePerBox",
      "toggleClick": "true",
      "showFlip": "false",
      "showUndo": "false",
      "shape": "circle 12 black red true",
      "cycleTimer": "tct"
    },
    { "name": "Feeder Count<br>(Fed another bot)",
      "code": "fo",
      "type": "counter"
    },
    { "name": "Others Fed # Pieces",
      "code": "of",
      "type": "counter"
    },
    { "name": "Dropped Pieces",
      "code": "dc",
      "type": "counter"
    }
  ],
  "endgame": [
    { "name": "Docking Timer",
      "code": "dt",
      "type": "timer"
    },
    { "name": "Final Status",
      "code": "fs",
      "type":"radio",
      "choices": {
        "park": "Parked<br>",
        "dock": "Docked (Not Engaged)<br>",
        "eng": "Engaged<br>",
        "att": "Attempted but failed<br>",
        "x": "Not attempted"
      },
      "defaultValue": "x"
    },
    { "name": "Total # of alliance<br>robots docked/engaged",
      "code": "dn",
      "type": "counter",
      "max": 3,
      "min": 0,
      "required": "true"
    }
  ],
  "postmatch": [
    { "name": "Driver Skill",
      "code": "ds",
      "type": "radio",
      "choices": {
        "less": "Less Effective<br>",
        "avg": "Average<br>",
        "more": "More Effective<br>",
        "x": "Not Observed"
      },
      "defaultValue": "x"
    },
    { "name": "Defense Rating",
      "code": "dr",
      "type": "radio",
      "choices": {
        "below": "Below Average<br>",
        "avg": "Average<br>",
        "good": "Good<br>",
        "exl": "Excellent<br>",
        "x": "Did not play defense"
      },
      "defaultValue": "x"
    },
    { "name": "Was Defended",
      "code": "wd",
      "type": "bool"
    },
    { "name": "Died/Immobilized",
      "code": "die",
      "type": "bool"
    },
    { "name": "Tippy<br>(almost tipped over)",
      "code": "tip",
      "type": "bool"
    },
    { "name": "Comments",
      "code": "co",
      "type": "text",
      "size": 15,
      "maxSize": 50
    }
  ]
}`;
