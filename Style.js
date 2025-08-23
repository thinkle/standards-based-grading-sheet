/* eslint-disable no-unused-vars */
/* Let's define some basics used for style throughout the project */

/* exported STYLE */
const STYLE = {
  // Font tokens
  FONT_FAMILY: 'Roboto',
  FONT_SIZE: '11',
  FONT_SIZE_SMALL: '8',
  FONT_SIZE_LARGE: '15',
  FONT_SIZE_XLARGE: '18',

  // Color tokens
  COLORS: {
    TEXT_COLOR: '#000000',
    INACTIVE_TEXT: '#333333',
    BRAND_PRIMARY: '#0033a0',
    BRAND_SECONDARY: '#464646',

    // Used for mastery gradient and neutral/text colors
    // MIN/MID/MAX are the endpoints/midpoint of the color scale (for backgrounds).
    // Picked to clearly convey bad → warning → good.
    // TEXT is the default foreground color to render on top of the scale.
    // BG kept for backward-compat (alias of TEXT).
    GRADE_SCALE: {
      MIN: '#c62828',            // red 800 (alert)
      MID: '#ef6c00',            // orange 800 (warning)
      MAX: '#2e7d32',            // green 800 (good)
      TEXT: '#ffffff',           // white text reads well on these tones
      BG: '#ffffff',             // alias: historical name, same as TEXT
      SCALE_MIN: '#c62828',
      SCALE_MAX: '#2e7d32'
    },

    // Per-level palettes; BG = regular tone, BG_BRIGHT = action tone for input areas
    LEVELS: {
      1: {
        BG: '#ffc37e', // soft orange
        BG_BRIGHT: '#ae5215', // strong orange
        TEXT: '#000000',
        TEXT_BRIGHT: '#ffffff'
      },
      2: {
        BG: '#ffe4d0', // soft yellow (fix 8-digit hex)
        BG_BRIGHT: '#ffd1af', // bright yellow
        TEXT: '#000000',
        TEXT_BRIGHT: '#19006a'
      },
      3: {
        BG: '#e0ffe8', // soft green (fix 8-digit hex)
        BG_BRIGHT: '#4f6337', // strong green (fix 8-digit hex)
        TEXT: '#000000',
        TEXT_BRIGHT: '#ffffff'
      },
      4: { // purple hues
        BG: '#e1bee7',
        BG_BRIGHT: '#ba68c8',
        TEXT: '#000000',
        TEXT_BRIGHT: '#ffffff'
      },
      5: { // gold for gold medal!
        BG: '#ffd700',
        BG_BRIGHT: '#fff494', // fix 8-digit hex
        TEXT: '#000000',
        TEXT_BRIGHT: '#ffffff'
      }
    },

    // General UI surfaces
    UI: {
      NEUTRAL_BG: '#f5f5f5',
      NEUTRAL_TEXT: '#333333',
      HEADER_BG: '#f0f3f5',
      HEADER_TEXT: '#000000',
      INPUT_HIGHLIGHT: '#fffbe6',
      INPUT_TEXT: '#000000',
      WARNING_BG: '#fff3cd',
      WARNING_TEXT: '#000000'
    }
  }
};



