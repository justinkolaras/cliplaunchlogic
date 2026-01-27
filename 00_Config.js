var CFG = {
  MAIN_NAMES: ['Main', 'MAIN'],
  URGENT_NAME: 'URGENT',
  COL_STATUS: 3,
  COL_INITIAL: 4,
  COL_FOLLOW_FLAG: 5,
  COL_FOLLOW_DATE: 6,
  COLORS: {
    ORANGE: '#ffa500',
    YELLOW: '#ffff00',
    GREEN: '#00ff00',
    RESET: '#ffffff',
    FLAG: '#00ffff'
  },
  TOKENS: {
    FLAG_NOTE: 'FLAG=1',
    ROW_ID_PREFIX: 'ROW_ID='
  }
};

var __NOTES_COL_CACHE = {};