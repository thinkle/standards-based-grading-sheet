/* eslint-disable no-unused-vars */
/* Let's define some basics used for style throughout the project */

/* !!! 
!!! NOTE: Using 8-digit HEX codes with Apps Script will cause
!!! WEIRD stuff to happen. Stick to 8 digits -- no alpha values! 
 !!! */

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

    // Per-level palettes; BG = regular tone, BG_ALT = striped alt; BG_BRIGHT = action tone for input areas, BG_BRIGHT_ALT = striped alt
    LEVELS: {
      1: {
        BG: '#fff1e6',        // very light peach
        BG_ALT: '#ffeadb',     // stripe alt
        BG_BRIGHT: '#ffd7ba',  // toned-down action
        BG_BRIGHT_ALT: '#ffcca8',
        TEXT: '#000000',
        TEXT_BRIGHT: '#000000',
        // Header-specific variants (used for unit/section header striping)
        BG_HEADER: '#ddcabd',           // default to BG_ALT
        TEXT_HEADER: '#000000',         // default to TEXT
        BG_BRIGHT_HEADER: '#e0b393',    // default to BG_BRIGHT_ALT
        BG_TEXT_HEADER: '#000000'       // default text color for bright headers
      },
      2: {
        BG: '#fff9e6',        // very light yellow
        BG_ALT: '#fff4d6',
        BG_BRIGHT: '#ffecb3',
        BG_BRIGHT_ALT: '#ffe39a',
        TEXT: '#000000',
        TEXT_BRIGHT: '#000000',
        BG_HEADER: '#d5ccb2',
        TEXT_HEADER: '#000000',
        BG_BRIGHT_HEADER: '#d9bf7e',
        BG_TEXT_HEADER: '#000000'
      },
      3: {
        BG: '#e8f5e9',        // very light green
        BG_ALT: '#ddf0df',
        BG_BRIGHT: '#c8e6c9',
        BG_BRIGHT_ALT: '#bde0be',
        TEXT: '#000000',
        TEXT_BRIGHT: '#000000',
        BG_HEADER: '#b4c5b6',
        TEXT_HEADER: '#000000',
        BG_BRIGHT_HEADER: '#849c85',
        BG_TEXT_HEADER: '#000000'
      },
      4: { // purple hues
        BG: '#f3e5f5',
        BG_ALT: '#eedeef',
        BG_BRIGHT: '#e1bee7',
        BG_BRIGHT_ALT: '#d7b2dd',
        TEXT: '#000000',
        TEXT_BRIGHT: '#000000',
        BG_HEADER: '#d3c3d4',
        TEXT_HEADER: '#000000',
        BG_BRIGHT_HEADER: '#b597b9',
        BG_TEXT_HEADER: '#000000'
      },
      5: { // gold for gold medal!
        BG: '#fff8e1',
        BG_ALT: '#fff2c6',
        BG_BRIGHT: '#ffe082',
        BG_BRIGHT_ALT: '#ffd460',
        TEXT: '#000000',
        TEXT_BRIGHT: '#000000',
        BG_HEADER: '#e3d7b1',
        TEXT_HEADER: '#000000',
        BG_BRIGHT_HEADER: '#e3be56',
        BG_TEXT_HEADER: '#000000'
      }
    },

    // General UI surfaces
    UI: {
      NEUTRAL_BG: '#f7f7f7',
      NEUTRAL_BG_ALT: '#f0f0f0',
      NEUTRAL_TEXT: '#333333',
      HEADER_BG: '#c9c9c9',
      HEADER_TEXT: '#000000',
      INPUT_HIGHLIGHT: '#fffdf0',
      INPUT_TEXT: '#000000',
      WARNING_BG: '#fff6da',
      WARNING_TEXT: '#000000',
      UNIT_TEXT_COLORS: [
        '#000000',
        '#B71C1C',
        '#0D47A1',
        '#1B5E20',
        '#880E4F',
        '#00695C',
        '#283593',
        '#4A148C'
      ]




    }
  }
};



