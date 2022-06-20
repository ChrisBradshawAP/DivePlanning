#define NFONTS 6
#define MAXSTR 100
//#define ASCENTRATE 9.00
#define ASCENTRATE 10.00
#define ASCENTRATEMEDIUM 12.00
#define ASCENTRATEFAST 18.00
//#define ASCENTDEPTHMEDIUM 9.00
//#define ASCENTDEPTHFAST 24.00
#define ASCENTDEPTHMEDIUM 499.00
#define ASCENTDEPTHFAST 624.00
#define ASCENTTIMEMEDIUM ((ASCENTDEPTHFAST-ASCENTDEPTHMEDIUM)/ASCENTRATEMEDIUM + ASCENTTIME)
#define ASCENTTIME ASCENTDEPTHMEDIUM/ASCENTRATE
#define DESCENTRATE 1000.00
#define DESCENTRATESLOW 20.00
#define SHAREWARE 0
#define AIR	1
#if SHAREWARE == 1
  #define NOX	0
  #define TRI  0
  #define REBX	0
  #define REB  0
#else
  #define NOX	1
  #define TRI  1
  #define REBX	1
  #define REB  1
#endif
#define DEBUGMODE 0
#define TRADEGAS 0
#define NUMGASMIX 10
#define OXYGENFRACTION (1.00 - nitrogenfraction - heliumfraction)
#define OXYGENFRACDECI (1.00 - nitrogenfracdec[i] - heliumfracdec[i])
#define OXYGENPOINTJI (1.00 - nitrogenpoint[j][i] - heliumpoint[j][i])
#define GASMIXO2 (1.00 - gasmix[j][k][0] - gasmix[j][k][1])
#define TIMEMOD 0.500
#define N2MIXPURE n2mix/(n2mix+hemix)
#define HEMIXPURE hemix/(n2mix+hemix)
#define O2MIXPOINTJI (1.00 - n2mixpoint[j][i] - hemixpoint[j][i])
#define ONE_POINT 1.001
#define VERSION "7.12C"
#define NUM_MICRO_STOPS 5
#define MICRO_STEP 12.00
#define SAFETYFACTOR_MICRO safetyfactor+micro_mode*18.00/500.00
