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
      "defaultValue": "2023alhu",
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
    { "name": "Starting Position",
      "code": "sp",
      "type": "radio",
      "choices": {
        "f": "Feeder Side<br>",
        "mi": "Middle<br>",
        "c": "Cable Protector"
      },
      "defaultValue": "f"
    }
  ],
  "auton": [
    { "name": "High Cones",
      "code": "ahc",
      "type": "counter",
      "max": 6,
      "min": 0
    },
    { "name": "High Cubes",
      "code": "ahcu",
      "type": "counter",
      "max": 3,
      "min": 0
    },
    { "name": "Middle Cones",
      "code": "amc",
      "type": "counter",
      "max": 6,
      "min": 0
    },
    { "name": "Middle Cubes",
      "code": "amcu",
      "type": "counter",
      "max": 3,
      "min": 0
    },
    { "name": "Low (Any Piece)",
      "code": "alc",
      "type": "counter",
      "max": 9,
      "min": 0
    },
    { "name": "Pieces Missed (Dropped)",
      "code": "pm",
      "type": "counter"
    },
    { "name": "Exited Community",
      "code": "ec",
      "type": "bool"
    },
    { "name": "Docked",
      "code": "ad",
      "type":"radio",
      "choices": {
        "d": "Docked (not Engaged)<br>",
        "e": "Engaged<br>",
        "a": "Attempted but failed<br>",
        "x": "Not attempted"
      },
      "defaultValue": "x"
    }
  ],
  "teleop": [
    { "name": "High Cones",
      "code": "hc",
      "type": "counter",
      "max": 6,
      "min": 0
    },
    { "name": "High Cubes",
      "code": "hcu",
      "type": "counter",
      "max": 3,
      "min": 0
    },
    { "name": "Middle Cones",
      "code": "mc",
      "type": "counter",
      "max": 6,
      "min": 0
    },
    { "name": "Middle Cubes",
      "code": "mcu",
      "type": "counter",
      "max": 3,
      "min": 0
    },
    { "name": "Low (Any Piece)",
      "code": "lc",
      "type": "counter",
      "max": 9,
      "min": 0
    },
    { "name": "Fouls",
      "code": "fl",
      "type": "counter"
    },
    { "name": "Tech Fouls",
      "code": "tf",
      "type": "counter"
    },
    { "name": "Yellow Card",
      "code": "yc",
      "type": "bool"
    },
    { "name": "Red Card",
      "code": "rc",
      "type": "bool"
    }
  ],
  "endgame": [
    { "name": "Final Status",
      "code": "fs",
      "type":"radio",
      "choices": {
        "d": "Docked (Not Engaged)<br>",
        "e": "Engaged<br>",
        "a": "Attempted but failed<br>",
        "x": "Not attempted"
      },
      "defaultValue": "x"
    },
    { "name": "Struggled to Engage",
      "code": "stg",
      "type": "bool"
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
        "b": "Below Average<br>",
        "a": "Average<br>",
        "aa": "Above Average<br>",
        "x": "Not Observed"
      },
      "defaultValue": "x"
    },
    { "name": "Defense Rating",
      "code": "dr",
      "type": "radio",
      "choices": {
        "b": "Below Average<br>",
        "a": "Average<br>",
        "aa": "Above Average<br>",
        "x": "Did not play defense"
      },
      "defaultValue": "x"
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
      "maxSize": 100
    }
  ]
}`;
