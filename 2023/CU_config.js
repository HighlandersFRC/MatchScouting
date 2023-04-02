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
        "1": "Feeder Side<br>",
        "2": "Middle<br>",
        "3": "Cable Protector"
      },
      "defaultValue": "1"
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
        "8": "Docked (not Engaged)<br>",
        "12": "Engaged<br>",
        "0": "Attempted but failed<br>",
        "-1": "Not attempted"
      },
      "defaultValue": "-1"
    }
  ],
  "teleop": [
    { "name": "Fouls",
      "code": "fl",
      "type": "counter"
    },
    { "name": "Tech Fouls",
      "code": "tf",
      "type": "counter"
    },
    { "name": "Yellow Card<br>Mark if Holding Card",
      "code": "yc",
      "type": "bool"
    },
    { "name": "Red Card<br>Mark if Holding Card",
      "code": "rc",
      "type": "bool"
    }
  ],
  "endgame": [
    { "name": "Final Status",
      "code": "fs",
      "type":"radio",
      "choices": {
        "6": "Docked (Not Engaged)<br>",
        "10": "Engaged<br>",
        "0": "Attempted but failed<br>",
        "-1": "Not attempted"
      },
      "defaultValue": "-1"
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
    { "name": "Driver Skill<br>(1-10)<br>(0=Not Observed)",
      "code": "ds",
      "type": "counter",
      "min": 0,
      "max": 10,
      "modifier": -1,
      "defaultValue": 0
    },
    { "name": "Defense Rating<br>(1-10)<br>(0=Did Not Defend)",
      "code": "dr",
      "type": "counter",
      "min": 0,
      "max": 10,
      "modifier": -1,
      "defaultValue": 0
    },
    { "name": "Died/Immobilized",
      "code": "die",
      "type": "bool"
    },
    { "name": "Tippy<br>(almost tipped over)",
      "code": "tip",
      "type": "bool"
    }
  ]
}`;
