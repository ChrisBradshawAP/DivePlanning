#include <stdio.h>
#include <math.h>
#include <stdlib.h>
#include <string.h>
#include <graph.h>
#include <conio.h>
#include <process.h>
#include <io.h>
#include <errno.h>
#include <dos.h>
#include <ctype.h>
#include <fcntl.h>
#include <sys\types.h>
#include <sys\utime.h>
#include <sys\stat.h>
#include <stddef.h>
#include <malloc.h>
#include <time.h>
#include "diveplan.h"

FILE _huge *fp2;

short action[5]  = { _GXOR,  _GXOR,_GPSET,	_GPRESET ,	   _GAND     };
char *descrip[5] = { "XOR   ", "XOR   " ,"XOR   ", "XOR   " , "OR    " };
char _huge *buffer, *bufferg, *buffert;
unsigned char divetitl[100], divetitllast[100];
char timebuf[10], datebuf[10], tradedatebuf[128], filebuf[20];

unsigned char licenseename[100], argvg[20], argvgn[3][20], divername[20][20];

#if TRADEGAS==1
  unsigned char *str[18] =
  {
  "Welcome to the", "Phoenix TRADE GAS PLANNER",
  "Copyright (c) Nick Bushell//Kevin Gurr 1994, 1995. All rights reserved","Serial number: ",
  "This software has been licensed to:", "Kevin Gurr",
  "The Licensee accepts that the Authors", "accept no liabilty for any gas fills calculated by this software","Read manual before use","Version Number: ","press any key to continue....","","WARNING",
  "ILLEGAL COPYING OF THIS SOFTWARE MAY RESULT IN","ERRONEOUS DECOMPRESSION COMPUTATION ERRORS",
  "IN THE ORIGINAL AND COPIED VERSION OF THE PROGRAM",""

  };

#else
  unsigned char *str[18] =
  {
  "Welcome to", "PRO-DIVE PLANNER",
  "Copyright(C) Nick Bushell\\Kevin Gurr 1992-1997. All rights reserved","Serial number: ",
  "This software has been licensed to:", "Kevin Gurr",
  "The Licensee accepts that the Authors", "accept no liabilty for any dives scheduled by this software","Read manual before use","Version Number: ","press any key to continue....","", "WARNING",
  "ILLEGAL COPYING OF THIS SOFTWARE MAY RESULT IN","ERRONEOUS DECOMPRESSION COMPUTATION ERRORS",
  "IN THE ORIGINAL AND COPIED VERSION OF THE PROGRAM",""
  };
#endif

unsigned char *str2[21] =
{
  "PRO-DIVE PLANNER MENU",
  "1  Air Dive Planner",
  "2  Nitrox Dive Planner",
  "3  Tri-mix Dive Planner",
  "4  Nitrox Re-Breather Dive Planner",
  "5  Tri-mix Re-Breather Dive Planner",
  "6  Metres",
  "7  Save Dive to disc",
  "8  Restore/Review Dive",
  "9  Safety Factor=",
  "A  Atmospheric=",
  "B  Breathing rate=",
  "C  Cylinder filling",
  "S  3 metre stop=",
  "G  Gas Tables \\ O  Optimization",
  "P  Laser Printer",
  "T  Time of Day",
  "M  Manual",
  "D  Pro-Log",
  "0  Exit",
  ""
};

unsigned char *str3[21] =
{
  "PRO-DIVE TRADE GAS FILL MENU",
  "L  View gas filling log",
  "1  Cylinder Gas details",
  "2  Gas costs",
  "3  Currency",
  "G  Dive Planner Menu",
  "P  Laser Printer",
  "M  Manual",
  "U  Bubbles",
  "0  Exit",
  "",
  "6  Metres",
  "7  Save Dive to disc",
  "8  Restore/Review Dive",
  "9  Safety Factor=",
  "A  Atmospheric=",
  "B  Breathing rate=",
  "C  Cylinder filling",
  "S  3 metre stop=",
  "T  Time of Day",
  "G  Trade gas filling",
};

double cnslooktable[55][2] = {
  0.500, -0.50,
  0.595,  0.00,
  0.635,  0.14,
  0.645,  0.15,
  0.665,  0.16,
  0.695,  0.17,
  0.725,  0.18,
  0.765,  0.20,
  0.785,  0.21,
  0.805,  0.22,
  0.855,  0.24,
  0.865,  0.25,
  0.885,  0.26,
  0.915,  0.28,
  0.935,  0.29,
  0.975,  0.31,
  1.005,  0.33,
  1.015,  0.34,
  1.045,  0.35,
  1.085,  0.40,
  1.105,  0.42,
  1.135,  0.43,
  1.165,  0.45,
  1.195,  0.47,
  1.235,  0.50,
  1.255,  0.51,
  1.295,  0.55,
  1.315,  0.56,
  1.345,  0.60,
  1.365,  0.61,
  1.375,  0.62,
  1.395,  0.64,
  1.400,  0.65,
  1.425,  0.68,
  1.440,  0.71,
  1.460,  0.74,
  1.465,  0.76,
  1.480,  0.78,
  1.495,  0.81,
  1.500,  0.83,
  1.520,  0.93,
  1.540,  1.04,
  1.560,  1.19,
  1.585,  1.47,
  1.605,  2.22,
  1.620,  5.00,
  1.650,  6.25,
  1.670,  7.69,
  1.700, 10.00,
  1.720, 12.50,
  1.740, 20.00,
  1.770, 25.00,
  1.780, 31.25,
  1.805, 50.00,
  2.500, 100.00

};

float stoplookup_factor[4][4] = {
  0.0, 1.0, 2.0, 3.0,
  0.0, 0.0, 2.0, 3.0,
  0.0, 1.5, 2.0, 3.0,
  0.0, .67, 1.34, 2.34
};

  short pixfact=1, piyfact=1, piyfactdec=0;
  /*unsigned char menunumber;*/
  unsigned char options[8];

  unsigned char safetytitlednum[10];
  unsigned char list[20];
  char fondir[_MAX_PATH];
  struct videoconfig vc;
  struct _fontinfo fi;
  short x, y, f;
  long prev_bk;
  short prev_fr, prev_cl;
  int xint, yint;

char cnumbuf[MAXSTR] = { MAXSTR + 2, 0 };
char tmpbuf[MAXSTR];
char  *numbuf;
char form[4], formlong[10], porb[4], cuftorltrmin[9], cuftorltr[9], cuftorltrnos[9], gascurrency[9];
unsigned char stitle[80];
short bubblex, bubbley = 30;

int air, n2, he, ppo2, ppi, ppf, number_gasses;
int x1graph1, y1graph1, x2graph1, y2graph1, x1graph2, y1graph2, x2graph2, y2graph2, xposgraph1, xposgraph2;
double hemixpointfracdec[110], n2mixpointfracdec[110];
int tissueno, depthno, deepeststop, numberstops, divenumber=0, numberpoints[10], divestart[10][3], divefinish[10][3], flytol[3];
long serialnumber, totaltimetosurface[10], timetofirststop[10];
double ppo2exptime_14[10], ppo2exptime_15[10], ppo2exptime_16[10], ppo2exptime_16plus[10], ppo2_now, ppo2fractionlast, ppo2cns[10], ppo2cnsmax[10], ppo2cnscurrent, ppo2cnsstart;
double tissue[16], a[16], b[16], halftime[16], tissuetemp[16], pambtoltiss[16], stoptimetissue[16], stoptime[110];
double tissuehe[16], ahe[16], bhe[16], halftimehe[16], tissuetemphe[16], pambtoltisshe[16], stoptimetissuehe[16],
  tissueorg[16], tissueorghe[16];
double absolutedepth, absolutedepthpure, depth, atmospheric=1.00, exposuretime=30.00, flytolmins, pambtolmin, pambtolmindepth, pigttolstopminus1[16], nitrogenfraction, nitrogenfracdec[110], heliumfraction, heliumfracdec[110];
int bailoutpointfracdec[110];
double depthpoint[10][110], timepointa[10][110], timepointb[10][110], timepointc[10][110], nitrogenpoint[10][110], heliumpoint[10][110], ppo2point[10][110], totaltimepoint[10][110], hemixpoint[10][110], n2mixpoint[10][110];
double ppo2fraction, ppo2fracdec[110], surftime, depthlast=0.00, currentstoppressure;
double nitrogenfractiondepth[10], exposuretimedepth[10], depthdepth[10], heliumfractiondepth[10], ppo2fractiondepth[10];
double hemixpointfractiondepth[10], n2mixpointfractiondepth[10];
int bailoutpointfractiondepth[10];
double stoptimeplus[10][110];
int bailoutpoint[10][110];
double depthmaxgraph, timetotalgraph, dailyotu[50], diveotu[10], missionotu, maxppo2[10], safetyfactor=0.00, dailyotustart[50];
double feetfactor=1.00, stopfactor=3.00, cuft_ltr_factor=28.316843, psifactor=1.00, maxdepthalarm, nitrogenfractioncalc, heliumfractioncalc;
/*double depthinc=0.75, decdepthinc=0.00, releaserate=0.9, ptolinc=0.10; Tom release */
double depthinc=1.00, decdepthinc=0.00, releaserate=1.00, ptolinc=0.00;
double nitrogenfractionlast, heliumfractionlast, fractionmax;
double gasmix[11][NUMGASMIX][3], gasmixbartime[10][NUMGASMIX], gasreservefraction[10][NUMGASMIX];
double gasmixtable[NUMGASMIX][3];
char gasstatus[NUMGASMIX], gasused[NUMGASMIX], set_gastable=0;
double filldive[10][NUMGASMIX], fillres[10][NUMGASMIX], filltotal[10][NUMGASMIX], cylindersize[10][NUMGASMIX], maxcylinderpressure[10][NUMGASMIX], freecylindersize[10][NUMGASMIX];
double breathingrate=10.0, breathingratedive, maxbreathratenumber=50.00, minbreathratenumber=5.00;
int  atmosphericdive, safetyfactordive, micro_factordive, menunumber, missionstart[3], missiontotal[3], bubblesmode=1, sixstopmode=0, lasermode=0, sixstopmodedive, tradegasscreenmode=0, trade=0;
double micro_mode=1.00, microfracdec[110];
short unsigned sprite[720];
double cdiff;
double atotal, btotal, nfrac, hefrac, tolstoppressure, currentfillpressure;
int toln2only;
double heprice, o2price, airprice, hecost, o2cost, aircost;
double tradecylindersize, tradefreecylindersize, trademaxcylinderpressure;
double airfillltr, hefillltr, o2fillltr, airfillprice, hefillprice, o2fillprice;
double airfillpricetotal, hefillpricetotal, o2fillpricetotal;
double airfillcosttotal, hefillcosttotal, o2fillcosttotal;
int numdivers;
int heo2display=1;
double oxygenfraction, oxygenfractionlast;
double hemix, n2mix, hemixlast, n2mixlast;
int bailout_breathable;
double tissuetrymag;
int bailout=0, bailoutlast=0;
double ppo2_limit_lower=1.00, ppo2_limit_upper=1.60;
int optimiseo2, automaticmode, autofinish, gas_table, repeat;
double missionotulast;
int missiontotallast[3];

void setinitial(void);
int tissueupdate(int temp_ascent);
void tissueorgtransfer(void);
void tissuetotissueorgtransfer(void);
void tissuetemptransfer(void);
void pambtolcalctrue(void);
void pambtolcalc(int temp_ascent);
double acalc(int i, int toln2only, int tempcalc);
double bcalc(int i, int toln2only, int tempcalc);
extern int getdecompressiontime(void);
void storedivedepthandtime();
extern int dispaydepthdata(void);
extern void currentcnsotudisplay(int x, int y);
void depthplot(void);
extern void plotdepthdata(int fullrun, int stopsprocessed, int depthprocessed);
int continuedivecheck(int airp, int n2p, int hep, int ppo2p);
void decomcalc(void);
extern void settitletext(void);
extern void setaxistext(void);
extern void setmicroaxistext(void);
extern void titleprofilegraph(int, double, double );
extern void borderdraw(void);
extern void printdivehistory(void);
extern void screenprint(int,int);
int licenseread(void);
int licensecheck(void);
void licensefeetmodewrite(void);
extern int getgas(int,int,double);
double stoptimetisscalc( int, int, double);
double tisscalc(int , double, int );
void ppo2print(double);
void setsafetyfactor(void);
void setatmospheric(void);
void setbreathingrate(void);
extern void setcylinderfill(void);
void setmissionstart(void);
extern void settitletextlarge(void);
int getdiscsaveddive(void);
void savediveondisc2(void);
void ppo2exposuretime(void);
void tissuegraph(void);
void divelist(void);
void prodivemenu(void);
long fileinfo( struct find_t *find );
void missiontotalupdate( double timeinc );
void flytolupdate(void);
void fractioncalcs(int);
void pubub(void);
void bubbles( void);
void abhalffile(void);
void drawbackground(void);
void backgroundtoggle( void);
void open_output( void);
void activate_graphic_mode( void);
void send_linefeed( void);
void restore_linefeed( void);
void delay1sec(void);
void activate_lasergraphic_mode( void);
void activate_lasergraphic_mode75( void);
void deactivate_lasergraphic_mode( void);
void laserleftmargin( void);
void lasernumberofbytes( int bytenum);
void laserposition( int , int );
void algorithmcalc(int i, double time, int fi);
void tissuetotissuetemptransfer(void);
extern double ascenttime (double deepest_depth);
extern double ascenttimediff (double deepest_depth, double depth);
extern void numgas_calc(int j, int i);
extern void printtoprinter(int j, int filetext);
extern void gasfillcalcs(
     int i,int j,int c,int k,int ii,int gppos, unsigned char title[100], double sftemp);
extern void gasfillsummary(
     int i,int j,int c,int k,int ii,int gppos, unsigned char title[100], double sftemp);
extern void printtradecylinderfill(
  int i, int c, int heinc, int topup, int entergaspercent,
  double sftemp, double fillpressure, double workingdepth, double workingppo2, double o2frac, double n2frac, double hefrac, double narcpressure, double hefill, double airfill, double o2fill, double n2fill,
  char message[MAXSTR],
  unsigned char title[100], unsigned char row,
  double hefillpressure, double o2fillpressure, double hefillfrac, double o2fillfrac, double n2fillfrac, double airtopoff,
  int logsave
  );
extern void tradecylinderfill(void);
extern void tradegasprice(void);
extern void tradegascurrency(void);
extern double fillgetgas(void);
extern void helpscreen(int);
extern int putdisc_gasdata(int row);
extern void getdisc_gasdata(void);
extern int getchyn( void);
extern int getchyn_defaulty( void);
extern char *cgetsn( char *buffer, char cn[10], char *default_string );
extern char *cgetsa( char *buffer );
extern void savefillcosts(void);
extern void graphicscreenprint( void);
extern void seto2optimise(void);
extern void setoxygenfraction(double ppo2depth);
extern int dispaydepthdatagas(void);
extern void putdisc_gasmixdata(char *mix_file);
extern void getdisc_gasmixdata(char *mix_file);
extern void tissupdate(int stopsprocessed);
extern void ttissupdate(int stopsprocessed);

char *timestr( unsigned t, char *buf );
char *datestr( unsigned d, char *buf );


FILE *fp, *fprn;




void main(int argc, char *argv[])

{
int licheck, i;
float d,dd;
char longtitle[100], numtitle[20];

  _dos_setfileattr( "diveplan.cop", _A_HIDDEN );

  strcpy(argvg,argv[1]);
  strcpy(argvgn[0],argv[1]);
  strcpy(argvgn[1],argv[2]);
  strcpy(argvgn[2],argv[3]);

  if( !strcmp(argv[1],"n2") || !strcmp(argv[2],"n2") )
       heo2display=0;
  else heo2display=1;
  if( !strcmp(argv[1],"hetol") || !strcmp(argv[2],"hetol") )
       toln2only=0;
  else toln2only=1;
  if( !strcmp(argv[1],"trade") || !strcmp(argv[1],"trade") ) {
    tradegasscreenmode=1;
    trade=1;
  }
  else trade=0;
  if(TRADEGAS>0) {
    trade=1;
    tradegasscreenmode=1;
    if(TRADEGAS==2) {
      tradegasscreenmode=0;
    }
  }
  if(!strcmp(argv[1],"bignose")) {
    printf("\nEnter depthinc(0.0-5.0):");
    scanf("%g",&d);
    printf("\nEnter decdepthinc(0.0-5.0):");
    scanf("%g",&dd);
    depthinc = (double)d;
    decdepthinc = (double)dd;
    printf("\nEnter releaserate(0.5-1.0):");
    scanf("%g",&d);
    releaserate = (double)d;
    printf("\nEnter pambtolinc(0.00-0.20):");
    scanf("%g",&d);
    ptolinc = (double)d;
    printf("\ndepthinc=%g decdepthinc=%g releaserate=%g ptolinc=%g",depthinc,decdepthinc,releaserate,ptolinc);
    getch();
  }
  strcpy(divetitl, "UNTITLED");
  strcpy(datebuf, "        ");
  strcpy( form , "msw");
  strcpy( formlong , "metres");
  strcpy ( porb , "bar");
  strcpy ( cuftorltrmin , "ltr/min");
  strcpy ( cuftorltr , "litres");
  strcpy ( cuftorltrnos , "litre");
  feetfactor=1.00; psifactor=1.00; stopfactor=3.00; cuft_ltr_factor=1.00;
  maxbreathratenumber=50.00, minbreathratenumber=5.00;
  strcpy(numtitle, divetitl);
  strcat(numtitle, ".mix");
  getdisc_gasmixdata(numtitle);

  getdisc_gasdata();

  /* Set highest available graphics mode and get configuration. */
  if( !_setvideomode( _MAXRESMODE)) exit( 1);
  _getvideoconfig( &vc);
  if(vc.numxpixels>500 && vc.numypixels>400) {
    vc.numxpixels=640;
    vc.numypixels=480;
  }
  /* vc.numcolors = 2; check */
  if( vc.numcolors > 2) {
    prev_bk = _setbkcolor( _BLUE);
    prev_cl = _setcolor(8);
  }
  _clearscreen( _GCLEARSCREEN);
  if( vc.numcolors > 2) prev_fr = _settextcolor(7);

  licenseread();
  /* Read header info from .FON files in current or given directory. */
  if( _registerfonts( "*.FON") <= 0) {
    _outtext( "Enter full path where .FON files are located: ");
    gets( fondir );
    strcat( fondir, "\\*.FON");
    if( _registerfonts( fondir) <= 0) {
      _outtext( "Error: can't register fonts");
      _setvideomode( _DEFAULTMODE);
      exit( 1);
    }
  }

  xint = (int)vc.numxpixels;
  yint = (int)vc.numypixels;
  if( xint < 500) pixfact = 2;
  if( yint < 400) piyfact = 2;
  if( yint < 220) piyfactdec = 6;

  /* Build options string. */
  settitletext();
  if( _setfont( list) >= 0) {
    do {
      _clearscreen( _GCLEARSCREEN);
      if( vc.numcolors > 2) {
	_setcolor(8);
	x = 60/pixfact; y = 100/piyfact;
	_rectangle( _GFILLINTERIOR, (x + 560)/pixfact, (y + 350)/piyfact, x, y );
	_setcolor(7);
	x = 40/pixfact; y = 80/piyfact;
	_rectangle( _GFILLINTERIOR, (x + 560)/pixfact, (y + 350)/piyfact, x, y );
      }
      licheck=0;
      for( f = 0; f < 18; f++) {
	if( f == 1 ) {
	  /* Rebuild options string. */
	  strcpy( options, "helv");
	  strcat( strcat( strcpy( list, "t'"), options), "'");
	  if(piyfact < 2) strcat( list, "h20w14b");
	  else strcat( list, "h10w7b");
	  if( _setfont( list) < 0 ) break;
	  if( vc.numcolors > 2) _setcolor( 1);
	}
	else {
	  /* Rebuild options string. */
	  settitletext();
	  if( vc.numcolors > 2) _setcolor( 1);
	}
	/* Use length of text to centeralise. */
	strcpy(longtitle, str[f]);
	if(f==3) {
	  ltoa(serialnumber, numtitle, 10 );
	  strcat(longtitle, numtitle );
	}
	if(f==5) strcpy(longtitle, licenseename);
	if(f==9) {
	  strcat(longtitle, VERSION );
	}
	x = (vc.numxpixels / 2) - (_getgtextextent(longtitle ) / 2);
	y = ( (480 / piyfact) / 5);
	/*y = (vc.numypixels / 4);*/
	if( _getfontinfo( &fi)) {
	  _outtext( "Error: Can't get font information");
	  break;
	}
	y += (f*20/piyfact);
	_moveto( x, y);

	/* display text. */
	_outgtext( longtitle);
      }
      if( licenseename[0] == 0x20) {
	licheck=1;
	if( !licensecheck() ) {
	  printf("No licensee name entered");
	  _setvideomode( _DEFAULTMODE);
	  exit(1);
	}
	_clearscreen( _GCLEARSCREEN);
	_setvideomode( _DEFAULTMODE);
	if(trade && tradegasscreenmode) system( "type manualg.txt | more" ); /* trade gas manual */
	else system( "type manual.txt | more" ); /* pro planner manual */
	printf("\n\n      Would you like a print out of this manual?  Y/<N> ");
	i = getchyn();
	printf("%c",i);
	if(i=='y' || i=='Y') {
	  printf("\n      Check printer is turned on\n");
	  printf(  "      Press any key to continue...\n");
	  getch();
	  if(trade && tradegasscreenmode) system( "print manualg.txt"); /* trade gas manual */
	  else system( "print manual.txt"); /* pro planner manual */
	  printf("      Press any key after print has finished...\n");
	  getch();
	}
	_setvideomode( _MAXRESMODE);
	if( vc.numcolors > 2) _setbkcolor( _BLUE);
	_clearscreen( _GCLEARSCREEN);
      }
    } while ( licheck ==1 );
  }
  else {
    _outtext( "Error: Can't set font: ");
    _outtext( list);
    }
  if(strcmp(argvgn[0], "pass") && strcmp(argvgn[0], "dlog")) getch();
  missionstart[0] = 1;
  missionstart[1] = 0;
  missionstart[2] = 0;
  breathingratedive = breathingrate;
  safetyfactordive = (int)(safetyfactor * 500.00);
  micro_factordive = (int)(micro_mode*100.00);
  atmosphericdive = (int)(atmospheric * 1000.00);
  divenumber=0;
  missionotu = 0.00;
  missiontotal[0] = missionstart[0];
  missiontotal[1] = missionstart[1];
  missiontotal[2] = missionstart[2];
  setinitial();
  abhalffile();
  pambtolcalc(0);

  while(1) {
    if(strcmp(argvgn[0], "pass")) prodivemenu();
    else {
      if(argvgn[1][0]) {
	if(getdiscsaveddive()) {
	  printdivehistory();
	  argvgn[1][0]=0;
	  argvgn[0][0]=0;
	  if(0) menunumber='2'+he;
	  else prodivemenu();
	}
	else {
	  argvgn[1][0]=0;
	  argvgn[0][0]=0;
	  strcpy(divetitl, "UNTITLED");
	}
      }
    }
    switch(menunumber) {

      case '1':
	#if AIR == 1
	  maxdepthalarm = 70.00;
	  if(!continuedivecheck( 1, 1, 0, 0 )) divenumber=0;
	  decomcalc();
	#endif
	break;

      case '2':
	#if NOX == 1
	  gas_table=1;
	  maxdepthalarm = 70.00;
	  if(!continuedivecheck( 0, 1, 0, 0 )) divenumber=0;
	  decomcalc();
	#endif
	break;

      case '3':
	#if TRI == 1
	  gas_table=1;
	  maxdepthalarm = 200.00;
	  if(!continuedivecheck( 0, 1, 1, 0 )) divenumber=0;
	  decomcalc();
	#endif
	break;

      case '4':
	#if REBX == 1
	  gas_table=1;
	  maxdepthalarm = 70.00;
	  if(!continuedivecheck( 0, 1, 0, 1 )) divenumber=0;
	  decomcalc();
	#endif
	break;

      case '5':
	#if REB == 1
	  gas_table=1;
	  maxdepthalarm = 250.00;
	  if(!continuedivecheck( 0, 1, 1, 1 )) divenumber=0;
	  decomcalc();
	#endif
	break;

      case '6':
	if(feetfactor == 1.00) {
	  feetfactor = 3.28084;
	  stopfactor = 3.048;
	  strcpy ( form , "fsw");
	  strcpy( formlong , "feet  ");
	  psifactor = 14.7;
	  strcpy ( porb , "psi");
	  strcpy ( cuftorltrmin , "cuft/min");
	  strcpy ( cuftorltr , "cuft  ");
	  strcpy ( cuftorltrnos , "cuft ");
	  maxbreathratenumber=2.00, minbreathratenumber=0.20;
	  cuft_ltr_factor=28.316843;
	}
	else {
	  feetfactor = 1.00;
	  stopfactor = 3.00;
	  strcpy ( form , "msw");
	  strcpy( formlong , "metres");
	  psifactor = 1;
	  strcpy ( porb , "bar");
	  strcpy ( cuftorltrmin , "ltr/min");
	  strcpy ( cuftorltr , "litres");
	  strcpy ( cuftorltrnos , "litre");
	  maxbreathratenumber=50.00, minbreathratenumber=5.00;
	  cuft_ltr_factor=1.00;
	}
	break;

      case '0':
	if( vc.numcolors > 2) {
	  _setbkcolor( prev_bk); /* restore original bkground */
	  _settextcolor(prev_fr);
	  _setcolor(prev_cl);
	}
	_setvideomode( _DEFAULTMODE);
	_unregisterfonts();
	licensefeetmodewrite();
	fcloseall();
	hfree( bufferg );
	hfree( buffert );
	return;
	break;

      case '8':
	if( !getdiscsaveddive() ) _strdate( datebuf );
	printdivehistory();
	break;

      case 'M':
	_clearscreen( _GCLEARSCREEN);
	_setvideomode( _DEFAULTMODE);
	if(trade && tradegasscreenmode) system( "type manualg.txt | more" ); /* trade gas manual */
	else system( "type manual.txt | more" ); /* pro planner manual */
	printf("\n\n      Would you like a print out of this manual?  Y/<N> ");
	i = getchyn();
	printf("%c",i);
	if(i=='y' || i=='Y') {
	  printf("\n      Check printer is turned on\n");
	  printf(  "      Press any key to continue...\n");
	  getch();
	  if(trade && tradegasscreenmode) system( "print manualg.txt"); /* trade gas manual */
	  else system( "print manual.txt" ); /* pro planner manual */
	  printf("      Press any key after print has finished...\n");
	  getch();
	}
	_setvideomode( _MAXRESMODE);
	if( vc.numcolors > 2) _setbkcolor( _BLUE);
	_clearscreen( _GCLEARSCREEN);
	break;

      case 'U':
	bubblesmode = bubblesmode ? 0 : 1 ;
	break;

#if SHAREWARE != 1
	case '7':
	if(!depthpoint[0][0]) {
	    if(!depthpoint[1][0]) break;
	}
	savediveondisc2();
	dispaydepthdatagas();
	break;

      case '9':
	setsafetyfactor();
	break;

      case 'A':
	setatmospheric();
	break;

      case 'B':
	setbreathingrate();
	break;

      case 'C':
	setcylinderfill();
	break;

      case 'L':
	_clearscreen( _GCLEARSCREEN);
	_setvideomode( _DEFAULTMODE);
	system( "type gasfill.txt | more" );
	printf("\n\n      Would you like a print out of this log?  Y/<N> ");
	i = getchyn();
	printf("%c",i);
	if(i=='y' || i=='Y') {
	  printf("\n      Check printer is turned on\n");
	  printf(  "      Press any key to continue...\n");
	  getch();
	  system( "print gasfill.txt" );
	  printf("      Press any key after print has finished...\n");
	  getch();
	}
	_setvideomode( _MAXRESMODE);
	if( vc.numcolors > 2) _setbkcolor( _BLUE);
	_clearscreen( _GCLEARSCREEN);
	break;

      case 'X':
	if(trade) tradecylinderfill();
	break;

      case 'Y':
	if(trade) tradegasprice();
	break;

      case 'Z':
	if(trade) tradegascurrency();
	break;

      case 'E':
	if(TRADEGAS==2) tradegasscreenmode = tradegasscreenmode ? 0 : 1 ;
	break;

      case 'S':
	sixstopmode = sixstopmode>1 ? 0 : sixstopmode+1 ;
	break;

      case 'T':
	setmissionstart();
	break;

      case 'P':
	lasermode = !lasermode ? 1 : lasermode==1 ? 2 : 0 ;
	break;

      case 'O':
	seto2optimise();
	dispaydepthdatagas();
	break;

      case 'G':
	dispaydepthdatagas();
	break;

      case 'D':
	if( vc.numcolors > 2) {
	  _setbkcolor( prev_bk); /* restore original bkground */
	  _settextcolor(prev_fr);
	  _setcolor(prev_cl);
	}
	_setvideomode( _DEFAULTMODE);
	_unregisterfonts();
	licensefeetmodewrite();
	fcloseall();
	hfree( bufferg );
	hfree( buffert );
	spawnl( P_OVERLAY, "divelog.exe", "divelog.exe", NULL);
	break;

      case 'Q':
	if( vc.numcolors > 2) {
	  _setbkcolor( prev_bk); /* restore original bkground */
	  _settextcolor(prev_fr);
	  _setcolor(prev_cl);
	}
	_setvideomode( _DEFAULTMODE);
	_unregisterfonts();
	licensefeetmodewrite();
	fcloseall();
	hfree( bufferg );
	hfree( buffert );
	return;
	break;

      case 27:
	if( vc.numcolors > 2) {
	  _setbkcolor( prev_bk); /* restore original bkground */
	  _settextcolor(prev_fr);
	  _setcolor(prev_cl);
	}
	_setvideomode( _DEFAULTMODE);
	_unregisterfonts();
	printf("\n\n\n\nNumber of x pixels = %d \nNumber of y pixels = %d", (int)vc.numxpixels, (int)vc.numypixels);
	printf("Returned number = %c", menunumber);
	printf("\nDecdepthinc = %f", decdepthinc);
	licensefeetmodewrite();
	fcloseall();
	hfree( bufferg );
	hfree( buffert );
	return;
	break;
#endif
    }
  }
  if( vc.numcolors > 2) {
    _setbkcolor( prev_bk); /* restore original bkground */
    _settextcolor(prev_fr);
    _setcolor(prev_cl);
  }
  _setvideomode( _DEFAULTMODE);
  printf("\n\n\n\nNumber of x pixels = %d \nNumber of y pixels = %d", (int)vc.numxpixels, (int)vc.numypixels);
  printf("Returned number = %c", menunumber);
}

void prodivemenu(void)
{
unsigned char title[80], titlednum[20];

  settitletext();
  _clearscreen( _GCLEARSCREEN);
  if( _setfont( list) >= 0) {
    if( vc.numcolors > 2) {
      _settextcolor(7);
      _setcolor( 8);
      x=12/pixfact; y=457/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 10)/pixfact, (y + 10)/piyfact, x/pixfact, y/piyfact );
      x=55/pixfact; y=430/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 20)/pixfact, (y + 20)/piyfact, x/pixfact, y/piyfact );
      x=105/pixfact; y=380/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 50)/pixfact, (y + 50)/piyfact, x/pixfact, y/piyfact );
      x=155/pixfact; y=355/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 75)/pixfact, (y + 75)/piyfact, x/pixfact, y/piyfact );
      x=155/pixfact; y=48/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 340)/pixfact, (y + 340)/piyfact, x/pixfact, y/piyfact );
      _setcolor( 7);
      x=10/pixfact; y=455/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 10)/pixfact, (y + 10)/piyfact, x/pixfact, y/piyfact );
      x=50/pixfact; y=425/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 20)/pixfact, (y + 20)/piyfact, x/pixfact, y/piyfact );
      x=100/pixfact; y=375/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 50)/pixfact, (y + 50)/piyfact, x/pixfact, y/piyfact );
      x=150/pixfact; y=350/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 75)/pixfact, (y + 75)/piyfact, x/pixfact, y/piyfact );
      x=150/pixfact; y=43/piyfact;
      _ellipse( _GFILLINTERIOR, (x + 340)/pixfact, (y + 340)/piyfact, x/pixfact, y/piyfact );
    }
    if(tradegasscreenmode) {
      for( f = 0; f <  10; f++) {
	if(TRADEGAS!=2 && f==5) f=6;
	if( vc.numcolors > 2) {
	  _setcolor( 0);
	  if( f == 0 ) _setcolor( 4);
	}
	/* Use length of text to centeralise. */
	strcpy( title , str3[f] );
	if(f==6) {
	  if(lasermode==1) strcpy( title ,"P  Laser Printer (small print)" );
	  if(lasermode==2) strcpy( title ,"P  Laser Printer (large print)" );
	  if(!lasermode)	 strcpy( title ,"P  Dot-matrix Printer");
	}
	x = (vc.numxpixels / 2) - (_getgtextextent( title ) / 2);
	y = ( (480 / piyfact) / 3);
	if( _getfontinfo( &fi)) {
	  _outtext( "Error: Can't get font information");
	  break;
	}
	y += (f*15/piyfact);
	_moveto( x, y);

	/* display text. */
	_outgtext( title);
      }
    }
    else {
      for( f = -1; f < 20; f++) {
	if( vc.numcolors > 2) {
	  _setcolor( 0);
	  if(SHAREWARE) if( (f!=6 && f!=1 && f!=8) && (f<17) ) _setcolor( 3);
	  if( f <= 0 ) _setcolor( 4);
	  #if AIR == 0
	    if( f == 1 ) _setcolor( 3);
	  #endif
	  #if NOX == 0
	    if( f == 2 ) _setcolor( 3);
	  #endif
	  #if TRI == 0
	    if( f == 3 ) _setcolor( 3);
	  #endif
	  #if REBX == 0
	    if( f == 4 ) _setcolor( 3);
	  #endif
	  #if REB == 0
	    if( f == 5 ) _setcolor( 3);
	  #endif
	}
	/* Use length of text to centeralise. */
	strcpy( title , str2[f] );
	if(f>=0) strcpy( title , str2[f] );
	else {
	  if(!strcmp("UNTITLED",divetitl)) strcpy(title, "");
	  else {
	    strcpy(title, "Dive: ");
	    strcat(title, divetitl);
	  }
	}
	if(f==6) {
	  if(feetfactor==1.00) strcpy( title , "6  Metres" );
	  else strcpy( title , "6  Feet");
	}
	if(f==9) {
	  itoa( ((int)(safetyfactor*500.00+0.49)), titlednum, 10);
	  strcat( title , titlednum);
	  strcat( title , "%");
	  itoa( ((int)(micro_mode*100.00+0.49)), titlednum, 10);
	  strcat( title , ", MicroBub ");
	  strcat( title , titlednum);
	  strcat( title , "%");
	}
	if(f==10) {
	  itoa( ((int)(atmospheric*1000.00+0.49)), titlednum, 10);
	  strcat( title , titlednum);
	  strcat( title , "mBar");
	}
	if(f==11) {
	  sprintf(titlednum, "%3.2f %s",feetfactor==1.00 ? breathingrate : breathingrate/cuft_ltr_factor, cuftorltrmin);
	  strcat( title , titlednum);
	}
	if(f==13) {
	  if(feetfactor==1.00) {
	    switch(sixstopmode) {
	      case 0:
		      strcpy( title ,"S  Last stop=3m" );
	      break;

	      case 1:
		      strcpy( title ,"S  Last stop=6m" );
	      break;

	      case 2:
		      strcpy( title ,"S  Last stops=6m & 4.5m" );
	      break;
	    }
	  }
	  else {
	    switch(sixstopmode) {
	      case 0:
		      strcpy( title ,"S  Last stop=10ft" );
	      break;

	      case 1:
		      strcpy( title ,"S  Last stop=20ft" );
	      break;

	      case 2:
		      strcpy( title ,"S  Last stops=20ft & 15ft" );
	      break;
	    }
	  }
	  /*
	  if(feetfactor==1.00) strcpy( title ,"S  3 metre stop=" );
	  else		     strcpy( title ,"S  10ft stop=");
	  if(sixstopmode) strcat( title , "OFF");
	    else  strcat( title , "ON");
	  */
	}
	if(f==15) {
	  if(lasermode==1) strcpy( title ,"P  Laser Printer (small print)" );
	  if(lasermode==2) strcpy( title ,"P  Laser Printer (large print)" );
	  if(!lasermode)	 strcpy( title ,"P  Dot-matrix Printer");
	}
	x = (vc.numxpixels / 2) - (_getgtextextent( title ) / 2);
	y = ( (480 / piyfact) / 5);
	/*y = (vc.numypixels / 5);*/
	if( _getfontinfo( &fi)) {
	  _outtext( "Error: Can't get font information");
	  break;
	}
	if(f<0) y += (f*28/piyfact);
	else y += (f*14/piyfact);
	_moveto( x, y);

	/* display text. */
	_outgtext( title);
      }
      if(trade) {
	setmicroaxistext();
	strcpy( title ,"G");
	x=180; y=365;
	_moveto( x, y);
	_outgtext( title);
	strcpy( title ,"Trade Gas");
	x=155; y=380;
	_moveto( x, y);
	_outgtext( title);
	strcpy( title ,"Filling");
	x=170; y=395;
	_moveto( x, y);
	_outgtext( title);
      }
    }
    if(!toln2only) {
      setmicroaxistext();
      strcpy( title ,"He/N2");
      x=105; y=390;
      _moveto( x, y);
      _outgtext( title);
      strcpy( title ,"TOL");
      x=110; y=405;
      _moveto( x, y);
      _outgtext( title);
    }
  } else {
    _outtext( "Error: Can't set font: ");
    _outtext( list);
    }

  if(tradegasscreenmode) {
    for(menunumber=0; menunumber != '0' && menunumber != 'M' && menunumber != 'P' && menunumber != 'U' && menunumber != 'Q' && menunumber != 27
	     && menunumber != 'X' && menunumber != 'Y' && menunumber != 'Z' && menunumber != 'G' && menunumber != 'L' ; ) {
       bubbles();
       menunumber = (getch());
       if(menunumber>0x30 && menunumber<0x34) menunumber += 0x27;
       if( menunumber > 0x39) menunumber = menunumber & 0xdf;
    }
  }
  else {
    for(menunumber=0; menunumber != '1' && menunumber != '2' && menunumber != '3' && menunumber != '4' && menunumber != '5' && menunumber != '6'
	&& menunumber != '7' && menunumber != '8' && menunumber != '9' && menunumber != '0' && menunumber != 'A'
	   && menunumber != 'B' && menunumber != 'C' && menunumber != 'S' && menunumber != 'T' && menunumber != 'M' && menunumber != 'P' && menunumber != 'U' && menunumber != 'Q' && menunumber != 27
	     && menunumber != 'X' && menunumber != 'Y' && menunumber != 'Z' && menunumber != 'G' && menunumber != 'D' && menunumber != 'L' && menunumber != 'O' && menunumber != 'E' ; ) {
       bubbles();
       menunumber = (getch());
       if( menunumber > 0x39) menunumber = menunumber & 0xdf;
    }
  }

}

int continuedivecheck(int airp, int n2p, int hep, int ppo2p)
{
int c, y=0;
  //if( (air>=airp && n2==n2p && he<=hep && ppo2==ppo2p && divenumber) || (ppo2p && hep) ) {
  if( (air>=airp && n2==n2p && he<=hep && ppo2<=ppo2p && divenumber) || (ppo2p && hep) ) {
    drawbackground();
  }
  bailout_breathable=0;
  //if( air>=airp && n2==n2p && he<=hep && ppo2==ppo2p && divenumber) {
  if( air>=airp && n2==n2p && he<=hep && ppo2<=ppo2p && divenumber) {
    helpscreen(26);
    if(!depthpoint[0][0]) {
      _settextposition( 8,20);
      sprintf(stitle, "VR3 TISSUE STATE LOADED FOR %s DIVE",divetitl);
      _outtext( stitle);
    }
    _settextposition( 10,10);
    sprintf(stitle, "Do you wish to continue %s mission sequence? Y/<N> ",divetitl);
    _outtext( stitle);
    c=getchyn();
    if(c=='y' || c=='Y') y=1;
  }

  air = airp;
  n2 = n2p;
  he = hep;
  ppo2 = ppo2p;
  return y;
}

void decomcalc(void)
{
int i, st, divenumberlast, j, c, oxtoolow;
double ddiff, dlast, breathingratedivelast;
int atmosphericdivelast, safetyfactordivelast, micro_factordivelast;

  breathingratedivelast = breathingratedive;
  safetyfactordivelast = safetyfactordive;
  micro_factordivelast = micro_factordive;
  atmosphericdivelast = atmosphericdive;
  missionotulast = missionotu;
  divenumberlast=divenumber;
  missiontotallast[0] = missiontotal[0];
  missiontotallast[1] = missiontotal[1];
  missiontotallast[2] = missiontotal[2];
  for(i=0;i<NUMGASMIX;i++) gasused[i]=0;
  if(divenumber) {
      divenumber--;
      numberpoints[divenumber]--;
      depthpoint[divenumber][numberpoints[divenumber]+1]=99999.00;
      timepointb[divenumber][numberpoints[divenumber]+1]=0.00;
      timepointc[divenumber][numberpoints[divenumber]+1]=0.00;
      do {
	oxtoolow=0;
	nitrogenfractionlast=0.79;
	heliumfractionlast=0.001;
	oxygenfractionlast=0.21;
	_settextposition( 12,10);
	sprintf(stitle, "Surface gas               ");
	_outtext( stitle) ;
	while(getgas( 12,22,0.00)<0);
	absolutedepth = atmospheric;
	absolutedepthpure = atmospheric;
	depth = 0.00;
	flytolmins = 1.00;

	for(j=0; j<16; j++) {
	  stoptimetissue[j]=0.00;
	  tolstoppressure=0.70;
	  pigttolstopminus1[j] = ( tolstoppressure ) / bcalc(j,toln2only,0) + acalc(j,toln2only,0);
	  fractioncalcs(0);
	  algorithmcalc(j, flytolmins, 0);
	  if( ( ((absolutedepth * heliumfractioncalc) < tissuetemphe[j]) || ((absolutedepth * nitrogenfractioncalc) < tissuetemp[j]) ) && (pigttolstopminus1[j] < (tissue[j] + tissuehe[j] ))) {
	  /* if( ( ((absolutedepth * heliumfractioncalc) < tissuehe[j]) || ((absolutedepth * nitrogenfractioncalc) < tissue[j]) ) && (pigttolstopminus1[j] < (tissue[j] + tissuehe[j] ))) { */
		stoptimetissue[j] = stoptimetisscalc( j, 0, flytolmins);
		if(stoptimetissue[j] > flytolmins)  {
		  flytolmins = stoptimetissue[j];
		}
		//if(stoptimetissue[j] == 29999.00) oxtoolow=1;
	  }
	}
	if(oxtoolow) {
	  _settextposition( 12,10);
	  sprintf(stitle, "WARNING: OXYGEN TOO LOW   ");
	  _outtext( stitle) ;
	  delay1sec();
	}
      } while (oxtoolow);
      helpscreen(7);
      _settextposition( 13,10);
      sprintf(stitle, "Surface time ____mins ");
      _outtext( stitle) ;
      _settextposition( 13,23);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      surftime = ((double)atoi( numbuf ));
      absolutedepthpure = atmospheric;
      depth = 0.00;
      exposuretime = surftime;

      numberpoints[divenumber]++;
      i = numberpoints[divenumber];
      depthpoint[divenumber][numberpoints[divenumber]+1]=99999.00;
      timepointb[divenumber][numberpoints[divenumber]+1]=0.00;
      timepointc[divenumber][numberpoints[divenumber]+1]=0.00;
      depthpoint[divenumber][i] = depth;
      timepointb[divenumber][i] = surftime;
      timepointc[divenumber][i] = 0.00;
      nitrogenpoint[divenumber][i] = nitrogenfraction;
      heliumpoint[divenumber][i] = heliumfraction;
      ppo2point[divenumber][i] = ppo2fraction;
      hemixpoint[divenumber][i] = hemix+0.0001;
      n2mixpoint[divenumber][i] = n2mix+0.0001;
      bailoutpoint[divenumber][i] = bailout;
      ppo2cnsstart = ppo2cnscurrent = ppo2cns[divenumber];
      /*
      ddiff = fabs(dlast - depthpoint[divenumber][i]);
      if( (dlast - depthpoint[divenumber][i]) < 0 ) timepointa[divenumber][i] = ddiff / DESCENTRATE;
      else timepointa[divenumber][i] = ascenttimediff(dlast, depthpoint[divenumber][i]);
      dlast = depthpoint[divenumber][i];
      */
      tissueupdate(0);
      divenumber++;
  }
  if(!divenumber) {
    strcpy(divetitl, "UNTITLED");
    strcpy(divetitllast, divetitl);

    breathingratedive = breathingrate;
    safetyfactordive = (int)(safetyfactor * 500.00);
    micro_factordive = (int)(micro_mode * 100.00);
    atmosphericdive = (int)(atmospheric * 1000.00);
    divenumber=0;
    breathingratedivelast = breathingratedive;
    safetyfactordivelast = safetyfactordive;
    micro_factordivelast = micro_factordive;
    atmosphericdivelast = atmosphericdive;
    missionotulast = missionotu;
    missionotu = 0.00;
    divenumberlast=divenumber;
    missiontotallast[0] = missiontotal[0];
    missiontotallast[1] = missiontotal[1];
    missiontotallast[2] = missiontotal[2];
    missiontotal[0] = missionstart[0];
    missiontotal[1] = missionstart[1];
    missiontotal[2] = missionstart[2];
    ppo2cnsstart = ppo2cnscurrent = 0.00;
    setinitial();
    abhalffile();
    pambtolcalc(0);
  }
  tissuetotissueorgtransfer();
  depthlast=0.00;
  ppo2fractionlast=ppo2_limit_upper;
  //bailout=bailoutlast=0;
  hemixlast=0.90;
  n2mixlast=0.00;
  nitrogenfractionlast=0.79;
  heliumfractionlast=0.001;
  oxygenfractionlast=0.21;
  do{
     do {
       bailout=bailoutlast=0;
       autofinish=automaticmode=0;
       tissueorgtransfer();
       ppo2fractionlast=ppo2_limit_upper;
       depthplot();
       if(dispaydepthdata()) {
	 if(!divenumber) {
	   strcpy(divetitl, divetitllast);
	   breathingratedive = breathingratedivelast;
	   safetyfactordive = safetyfactordivelast;
	   micro_factordive = micro_factordivelast;
	   atmosphericdive = atmosphericdivelast;
	   missiontotal[0] = missiontotallast[0];
	   missiontotal[1] = missiontotallast[1];
	   missiontotal[2] = missiontotallast[2];
	   missionotu = missionotulast;
	   divenumber = divenumberlast;
	 }
	 return;
       }

     } while ( !getdecompressiontime() ) ;
       plotdepthdata(1, numberstops, 5);
       if(surftime) {
	 divenumber++;
	 tissuetotissueorgtransfer();
       }
       else {
	 helpscreen(59);
	 _settextposition( 11,3);
	 sprintf(stitle, "Re-Edit dive ? Y/<N> ");
	 _outtext( stitle);
	 c=getchyn();
	 if(c=='y' || c=='Y') {
	   tissueorgtransfer();
	   y=1;
	 } else y=0;
       }


   }while(surftime && (divenumber < 10) || y);
   divenumber++;

}

void setinitial(void)
{
  int i,j;

/*  a[0]=1.900;
  a[1]=1.450;
  a[2]=1.030;
  a[3]=0.882;
  a[4]=0.717;
  a[5]=0.575;
  a[6]=0.468;
  a[7]=0.441;
  a[8]=0.415;
  a[9]=0.416;
  a[10]=0.369;
  a[11]=0.369;
  a[12]=0.255;
  a[13]=0.255;
  a[14]=0.255;
  a[15]=0.255;

  b[0]=0.800;
  b[1]=0.800;
  b[2]=0.800;
  b[3]=0.826;
  b[4]=0.845;
  b[5]=0.860;
  b[6]=0.870;
  b[7]=0.903;
  b[8]=0.908;
  b[9]=0.939;
  b[10]=0.946;
  b[11]=0.946;
  b[12]=0.962;
  b[13]=0.962;
  b[14]=0.962;
  b[15]=0.962;
*/
  a[0]=1.2599;
  a[1]=1.0000;
  a[2]=0.8618;
  a[3]=0.7562;
  a[4]=0.6667;
  a[5]=0.5600;
  a[6]=0.4947;
  a[7]=0.45;
  a[8]=0.4187;
  a[9]=0.3798;
  a[10]=0.3497;
  a[11]=0.3223;
  a[12]=0.2850;
  a[13]=0.2737;
  a[14]=0.2523;
  a[15]=0.2327;

  b[0]=0.5050;
  b[1]=0.6514;
  b[2]=0.7222;
  b[3]=0.7825;
  b[4]=0.8126;
  b[5]=0.8434;
  b[6]=0.8693;
  b[7]=0.8910;
  b[8]=0.9092;
  b[9]=0.9222;
  b[10]=0.9319;
  b[11]=0.9403;
  b[12]=0.9477;
  b[13]=0.9544;
  b[14]=0.9602;
  b[15]=0.9653;


  halftime[0]=4.00;
  halftime[1]=8.00;
  halftime[2]=12.50;
  halftime[3]=18.50;
  halftime[4]=27.00;
  halftime[5]=38.30;
  halftime[6]=54.30;
  halftime[7]=77.00;
  halftime[8]=109.0;
  halftime[9]=146.0;
  halftime[10]=187.0;
  halftime[11]=239.0;
  halftime[12]=305.0;
  halftime[13]=390.0;
  halftime[14]=498.0;
  halftime[15]=635.0;


  ahe[0]=1.7424;
  ahe[1]=1.383;
  ahe[2]=1.1919;
  ahe[3]=1.0458;
  ahe[4]=0.922;
  ahe[5]=0.8205;
  ahe[6]=0.7305;
  ahe[7]=0.6502;
  ahe[8]=0.595;
  ahe[9]=0.5545;
  ahe[10]=0.5333;
  ahe[11]=0.5189;
  ahe[12]=0.5181;
  ahe[13]=0.5176;
  ahe[14]=0.5172;
  ahe[15]=0.5119;

  bhe[0]=0.4245;
  bhe[1]=0.5747;
  bhe[2]=0.6527;
  bhe[3]=0.7223;
  bhe[4]=0.7582;
  bhe[5]=0.7957;
  bhe[6]=0.8279;
  bhe[7]=0.8553;
  bhe[8]=0.8757;
  bhe[9]=0.8903;
  bhe[10]=0.8997;
  bhe[11]=0.9073;
  bhe[12]=0.9122;
  bhe[13]=0.9171;
  bhe[14]=0.9217;
  bhe[15]=0.9267;

  halftimehe[0]=1.51;
  halftimehe[1]=3.02;
  halftimehe[2]=4.72;
  halftimehe[3]=6.99;
  halftimehe[4]=10.21;
  halftimehe[5]=14.48;
  halftimehe[6]=20.53;
  halftimehe[7]=29.11;
  halftimehe[8]=41.2;
  halftimehe[9]=55.19;
  halftimehe[10]=70.69;
  halftimehe[11]=90.34;
  halftimehe[12]=115.29;
  halftimehe[13]=147.42;
  halftimehe[14]=188.24;
  halftimehe[15]=240.03;
/*
  ahe[0]=1.2599;
  ahe[1]=1.0000;
  ahe[2]=0.8618;
  ahe[3]=0.7562;
  ahe[4]=0.6667;
  ahe[5]=0.5600;
  ahe[6]=0.4947;
  ahe[7]=0.45;
  ahe[8]=0.4187;
  ahe[9]=0.3798;
  ahe[10]=0.3497;
  ahe[11]=0.3223;
  ahe[12]=0.2850;
  ahe[13]=0.2737;
  ahe[14]=0.2523;
  ahe[15]=0.2327;

  bhe[0]=0.5050;
  bhe[1]=0.6514;
  bhe[2]=0.7222;
  bhe[3]=0.7825;
  bhe[4]=0.8126;
  bhe[5]=0.8434;
  bhe[6]=0.8693;
  bhe[7]=0.8910;
  bhe[8]=0.9092;
  bhe[9]=0.9222;
  bhe[10]=0.9319;
  bhe[11]=0.9403;
  bhe[12]=0.9477;
  bhe[13]=0.9544;
  bhe[14]=0.9602;
  bhe[15]=0.9653;

  halftimehe[0]=1.51;
  halftimehe[1]=3.02;
  halftimehe[2]=4.72;
  halftimehe[3]=6.99;
  halftimehe[4]=10.21;
  halftimehe[5]=14.48;
  halftimehe[6]=20.53;
  halftimehe[7]=29.11;
  halftimehe[8]=41.20;
  halftimehe[9]=55.19;
  halftimehe[10]=70.69;
  halftimehe[11]=90.34;
  halftimehe[12]=115.29;
  halftimehe[13]=147.42;
  halftimehe[14]=188.24;
  halftimehe[15]=240.03;
*/

  for(i=0; i<16; i++) {
    tissue[i]=0.79*atmospheric;
    tissuehe[i]=0.00;
  }
  nitrogenfraction = 0.79;
  heliumfraction = 0.001;

  x1graph1=60; y1graph1=220-piyfactdec; x2graph1=315; y2graph1=375-piyfactdec; xposgraph1=0;
  x1graph2=355; y1graph2=220-piyfactdec; x2graph2=610; y2graph2=375-piyfactdec; xposgraph2=295;

  for(i=0; i<10; i++) {
    for(j=0; j<NUMGASMIX; j++) {
      gasmix[i][j][0]=0.00;
      gasmix[i][j][1]=0.00;
      gasmixbartime[i][j]=0.00;
      gasreservefraction[i][j]=0.00;
      filldive[i][j]=0.00;
      fillres[i][j]=0.00;
      filltotal[i][j]=0.00;
      cylindersize[i][j]=0.00;
      maxcylinderpressure[i][j]=0.00;
      freecylindersize[i][j]=0.00;
    }
  }


}

int tissueupdate(int temp_ascent)
{
int i, j;
double depthdiff, exposuretimelast, otuadd, depthdiff2;

  if(!exposuretime) return 0; /* return if exposure time 0.00 */

  depthdiff = fabs(depthlast - depth);
  depthdiff2 = depthdiff/2.00;
  if(depth>depthlast) absolutedepth = (depthdiff2+depthlast)/10.00 + atmospheric ;
  else		      absolutedepth = (depthdiff2+depth)/10.00 + atmospheric ;
  exposuretimelast = exposuretime;
  if( (depthlast - depth) < 0.00 ) exposuretime = depthdiff / DESCENTRATE;
  else				exposuretime = ascenttimediff(depthlast, depth);
  _settextposition(1,1);
  if(!strcmp(argvg,"bignose4")) {
    printf("Absdepth=%g, exptime=%g  ",absolutedepth,exposuretime);
  }

  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("\n\n\n\n\n\nA1bsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
  if(ppo2) fractioncalcs((int)depth); // Use previous values unless rebreather
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("A2bsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
  if(!temp_ascent && (( 2.00 * absolutedepth * ( ONE_POINT - nitrogenfraction - heliumfraction) - 1.00) > 0) ) {
    otuadd = exposuretime * pow( (2.00 * absolutedepth * ( ONE_POINT - nitrogenfraction - heliumfraction) - 1.00),	0.833);
    j = divestart[divenumber][0];
    diveotu[divenumber] = diveotu[divenumber] + otuadd ;
    dailyotu[j] = dailyotu[j] + otuadd ;
    missionotu = missionotu + otuadd ;
  }
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B2aBbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
  for(i=0; i<16; i++) {

    algorithmcalc(i, exposuretime, (int)depth);

  }
  if(temp_ascent) {
    pambtolcalc(1);
    return 0;
  }
  if(!temp_ascent) {
    missiontotalupdate(TIMEMOD);
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B2bBbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    ppo2exposuretime();
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B2cBbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    tissuetemptransfer();
    absolutedepth= (depth/10.00) + atmospheric ;
    depthlast = depth;
    exposuretime = exposuretimelast; /* To keep depth and exposuretime the same as when entered*/
    if(!strcmp(argvg,"bignose4")) {
      printf("\nAbsdepth=%g, exptime=%4g",absolutedepth,exposuretime);
      printf("\nd=%3g decd=%3g rr=%3g pt=%3g",depthinc,decdepthinc,releaserate,ptolinc);
    }
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B2dBbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    fractioncalcs((int)depth);
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B2eBbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    if( maxppo2[divenumber] < (absolutedepthpure * ( ONE_POINT - nitrogenfraction - heliumfraction)) )	maxppo2[divenumber] = absolutedepthpure * ( 1.00 - nitrogenfraction - heliumfraction);
    if(( 2.00 * absolutedepth * ( ONE_POINT - nitrogenfraction - heliumfraction) - 1.00) > 0) {
      otuadd = exposuretime * pow( (2.00 * absolutedepth * ( ONE_POINT - nitrogenfraction - heliumfraction) - 1.00),	0.833);
      j = divestart[divenumber][0];
      diveotu[divenumber] = diveotu[divenumber] + otuadd ;
      dailyotu[j] = dailyotu[j] + otuadd ;
      missionotu = missionotu + otuadd ;
    }
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
      printf("B3Bbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
    }

    for(i=0; i<16; i++) {
      if(!strcmp(argvg,"bignose")) {
	(cdiff=(nitrogenfractioncalc * absolutedepth - tissue[i] ));
	printf("\ncdiffn2=%g", cdiff);
	(cdiff=(heliumfractioncalc * absolutedepth - tissuehe[i] ));
	printf(" cdiffhe=%g", cdiff);
      }
      algorithmcalc(i, exposuretime, (int)depth);
      if(!strcmp(argvg,"bignose")) {
	printf("ct=%g, cthe=%g, nf=%g, hf=%g",tissuetemp[i],tissuetemphe[i],nitrogenfraction,heliumfraction);
      }
    }

    if(!strcmp(argvg,"bignose")) {
      while(!kbhit);
      getch();
    }
    missiontotalupdate(0.00);
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B4Bbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    ppo2exposuretime();
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B5Bbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    tissuetemptransfer();
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B6Bbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    pambtolcalc(0);
  if(divenumber==1 && !strcmp(argvg,"bignose9")) {
    printf("B7Bbsdepth=%g, exptime=%g, he=%g  ",absolutedepthpure,exposuretime,heliumfractioncalc);
  }
    return 1;
  }
}

void fractioncalcs(int fpo2depth)
{
  if(fpo2depth==12 && divenumber && !strcmp(argvg,"bignoseb")) {
    printf("1fp=%d, pp%g, ab%g, n2%g, he%g",fpo2depth, ppo2fraction, absolutedepthpure, n2mix, hemix );
    while(!kbhit());
    getch();
  }
  if(ppo2 && fpo2depth && !bailout) {
    if(n2) {
      heliumfraction = 0.001;
      if((absolutedepthpure - ppo2fraction - 0.10) < 0) {
	nitrogenfraction = 0.05 ;
	ppo2fraction = absolutedepthpure * (1.00 - nitrogenfraction) ;
      }
      else nitrogenfraction = (absolutedepthpure - ppo2fraction) / absolutedepthpure;
    }
    if(he) {
      //nitrogenfraction = 0.001;
      if((absolutedepthpure - ppo2fraction - 0.10) < 0) {
	heliumfraction = HEMIXPURE * 0.05 ;
	nitrogenfraction = N2MIXPURE * 0.05 + 0.001;
	ppo2fraction = absolutedepthpure * (1.00 - heliumfraction - nitrogenfraction) ;
      }
      else {
	heliumfraction = (HEMIXPURE)*(absolutedepthpure - ppo2fraction) / absolutedepthpure;
	nitrogenfraction = (N2MIXPURE)*(absolutedepthpure - ppo2fraction) / absolutedepthpure;
      }
    }
  }

  if( (he) ) {
    nitrogenfractioncalc = nitrogenfraction + ( fpo2depth ? SAFETYFACTOR_MICRO : 0 ) / 2.00 ;
    heliumfractioncalc = heliumfraction + ( fpo2depth ? SAFETYFACTOR_MICRO : 0 ) / 2.00 ;
  }
  else {
      nitrogenfractioncalc = nitrogenfraction + ( fpo2depth ? SAFETYFACTOR_MICRO : 0 ) ;
      heliumfractioncalc = 0.00 ;
  }

  if(divenumber && !strcmp(argvg,"bignoseb")) {
    printf("nf=%g, hf=%g, nfc=%g, hfc=%g",nitrogenfraction,heliumfraction,nitrogenfractioncalc,heliumfractioncalc);
    while(!kbhit());
    getch();
  }
}

void tissuetotissueorgtransfer(void)
{
int i;
  for(i=0; i<16; i++) {
    tissueorg[i] = tissue[i];
    tissueorghe[i] = tissuehe[i];
  }
  missiontotallast[0] = missiontotal[0];
  missiontotallast[1] = missiontotal[1];
  missiontotallast[2] = missiontotal[2];
  missionotulast = missionotu;
  ppo2cnsstart = ppo2cnscurrent;
  for(i=0; i<50; i++) {
    dailyotustart[i] = dailyotu[i];
  }
}

void tissueorgtransfer(void)
{
int i;
  for(i=0; i<16; i++) {
    tissue[i] = tissueorg[i];
    tissuehe[i] = tissueorghe[i];
  }
  missiontotal[0] = missiontotallast[0];
  missiontotal[1] = missiontotallast[1];
  missiontotal[2] = missiontotallast[2];
  missionotu = missionotulast;
  ppo2cnscurrent = ppo2cnsstart;
  diveotu[divenumber] = 0.00;
  for(i=0; i<50; i++) {
    dailyotu[i] = dailyotustart[i];
  }
}

void tissuetotissuetemptransfer(void)
{
int i;
for(i=0; i<16; i++) {
  tissuetemp[i] = tissue[i];
  tissuetemphe[i] = tissuehe[i];
  }
}

void tissuetemptransfer(void)
{
int i;
for(i=0; i<16; i++) {
  tissue[i] = tissuetemp[i];
  tissuehe[i] = tissuetemphe[i];
  }
}

void pambtolcalctrue()
{
int i;

  exposuretime = 0.001;
  for(i=0, depth=0.00; i<10 ;i++, depth=pambtolmindepth*1.03) {
    tissueupdate(1);
    if( (pambtolmindepth-depth)<0.05) break;
  }
  pambtolmindepth++;
}

void pambtolcalc(int temp_ascent)
{
int i;

  pambtolmin=0.00;
  pambtolmindepth=0.00;
  for(i=0; i<16; i++) {
    if(temp_ascent) pambtoltiss[i] = bcalc(i,toln2only,temp_ascent) * ( tissuetemp[i] + tissuetemphe[i] - acalc(i,toln2only,temp_ascent) );
    else	    pambtoltiss[i] = bcalc(i,toln2only,temp_ascent) * ( tissue[i] + tissuehe[i] - acalc(i,toln2only,temp_ascent) );
    if(pambtoltiss[i] > pambtolmin) {
      pambtolmin = pambtoltiss[i] + ptolinc;
      pambtolmindepth = (pambtolmin - atmospheric) * 10.00;
      tissueno=i;
    }
    //if(!strcmp(argvg,"bignose9")&&temp_ascent) {
    if(!strcmp(argvg,"bignose9")) {
	printf("\nPambtoltiss[%d]=%4g, depthlast=%4g ",i,pambtoltiss[i],depthlast);
    }
  }
}

double acalc(int i, int toln2only, int tempcalc)
{
    if(toln2only) return a[i];

    if(tempcalc) {
      nfrac=tissuetemp[i]/(tissuetemp[i] + tissuetemphe[i]);
      hefrac=tissuetemphe[i]/(tissuetemp[i] + tissuetemphe[i]);
    }
    else {
      nfrac=tissue[i]/(tissue[i] + tissuehe[i]);
      hefrac=tissuehe[i]/(tissue[i] + tissuehe[i]);
    }
    if(!hefrac) hefrac=0.005;
    if(!nfrac) nfrac=0.005;
    /*
    atotal = (nitrogenfraction * a[i] + heliumfraction * ahe[i])
				/ (nitrogenfraction + heliumfraction);
    btotal = (nitrogenfraction * b[i] + heliumfraction * bhe[i])
				/ (nitrogenfraction + heliumfraction);
    return a[i];
    */
    atotal = (nfrac * a[i] + hefrac * ahe[i])
				/ (nfrac + hefrac);
    btotal = (nfrac * b[i] + hefrac * bhe[i])
				/ (nfrac + hefrac);
    return atotal;

}

double bcalc(int i, int toln2only, int tempcalc)
{
    if(toln2only) return b[i];

    if(tempcalc) {
      nfrac=tissuetemp[i]/(tissuetemp[i] + tissuetemphe[i]);
      hefrac=tissuetemphe[i]/(tissuetemp[i] + tissuetemphe[i]);
    }
    else {
      nfrac=tissue[i]/(tissue[i] + tissuehe[i]);
      hefrac=tissuehe[i]/(tissue[i] + tissuehe[i]);
    }
    if(!nfrac) nfrac=0.005;
    if(!hefrac) hefrac=0.005;
    /*
    atotal = (nitrogenfraction * a[i] + heliumfraction * ahe[i])
				/ (nitrogenfraction + heliumfraction);
    btotal = (nitrogenfraction * b[i] + heliumfraction * bhe[i])
				/ (nitrogenfraction + heliumfraction);
    return b[i];
    */
    atotal = (nfrac * a[i] + hefrac * ahe[i])
				/ (nfrac + hefrac);
    btotal = (nfrac * b[i] + hefrac * bhe[i])
				/ (nfrac + hefrac);
    return btotal;
}

void depthplot(void)
{
/* PLOT depth against time GRAPH */
int numtime, numdepth, snum, i, j, x, y, p;
double depthmax=0.00, timetotal=0.00, ddiff, dlast, tlast, timepointtotal=0;
unsigned char title[100], titlednum[10];
long imsize;

    _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact);

  /* Set highest available graphics mode and get configuration. */
    if( vc.numcolors > 2)_setbkcolor( _BLACK);
    _clearscreen( _GCLEARSCREEN);
    if( vc.numcolors > 2) _setbkcolor( _BLUE);
    if( vc.numcolors > 2) _setcolor(7);
    if( vc.numcolors > 2) _settextcolor(7);

    if( vc.numcolors > 2) {
    _setlinestyle( 0xffff);

    _setcolor(7);
     x=0; y=0;
     _rectangle( _GFILLINTERIOR, vc.numxpixels, vc.numypixels, x, y );

    _setcolor(0);
     x=8/pixfact; y=46/piyfact;
     _rectangle( _GFILLINTERIOR, (vc.numxpixels-8)/pixfact, 178/piyfact, x, y );

    _setcolor(15);
    _moveto_w( (double)vc.numxpixels-5.00/(double)pixfact, 5.00/(double)pixfact);
      _lineto_w( 5.00/(double)pixfact, 5.00/(double)piyfact );
      _lineto_w( 5.00/(double)pixfact, (double)(480/piyfact)-5.00/(double)piyfact );
    _setcolor(1);
      _lineto_w( (double)vc.numxpixels-5.00/(double)pixfact, (double)(480/piyfact)-5.00/(double)piyfact );
      _lineto_w( (double)vc.numxpixels-5.00/(double)pixfact, 5.00/(double)piyfact );

    //_setcolor(1);
    //_moveto_w( (double)vc.numxpixels-11.00/(double)pixfact, 45.00/(double)pixfact);
    //	_lineto_w( 11.00/(double)pixfact, 45.00/(double)piyfact );

    _setcolor(7);
     x=11/pixfact; y=11/piyfact;
     _rectangle( _GFILLINTERIOR, vc.numxpixels-11/pixfact, 44/piyfact, x, y );

    _setcolor(7);
     x=20/pixfact; y=180/piyfact;
     _rectangle( _GFILLINTERIOR, vc.numxpixels-x, (480/piyfact)-20/piyfact, x, y );

    _setcolor(8);
     x=(x1graph2+6)/pixfact; y=y1graph2/piyfact;
     _rectangle( _GFILLINTERIOR, x+(x2graph2-x1graph2-6)/pixfact, y+(y2graph2-y1graph2-6)/piyfact, x, y );
/*    imsize = _imagesize( x, y, x+(x2graph2-x1graph2-6)/pixfact, y+(y2graph2-y1graph2-6)/piyfact  );
    bufferg = halloc( imsize, 1 );
    _getimage( x, y, x+(x2graph2-x1graph2-6)/pixfact, y+(y2graph2-y1graph2-6)/piyfact, bufferg );
*/

    _setcolor(8);
     x=(x1graph1+6)/pixfact; y=y1graph1/piyfact;
     _rectangle( _GFILLINTERIOR, x+(x2graph1-x1graph1-6)/pixfact, y+(y2graph1-y1graph1-6)/piyfact, x, y );

    _setcolor(15);
    _moveto_w( (double)vc.numxpixels-20.00/(double)pixfact, 180.00/(double)pixfact);
      _lineto_w( 20.00/(double)pixfact, 180.00/(double)piyfact );
      _lineto_w( 20.00/(double)pixfact, (double)(480/piyfact)-20.00/(double)piyfact );
    _setcolor(1);
      _lineto_w( (double)vc.numxpixels-20.00/(double)pixfact, (double)(480/piyfact)-20.00/(double)piyfact );
      _lineto_w( (double)vc.numxpixels-20.00/(double)pixfact, 180.00/(double)piyfact );

    _setcolor(15);
    _moveto_w( 310.00/(double)pixfact, 190.00/(double)pixfact);
      _lineto_w( 70.00/(double)pixfact, 190.00/(double)piyfact );
      _lineto_w( 70.00/(double)pixfact, 215.00/(double)piyfact );
    _setcolor(1);
      _lineto_w( 310.00/(double)pixfact, 215.00/(double)piyfact );
      _lineto_w( 310.00/(double)pixfact, 190.00/(double)piyfact );

    _setcolor(15);
    _moveto_w( 610.00/(double)pixfact, 190.00/(double)pixfact);
      _lineto_w( 360.00/(double)pixfact, 190.00/(double)piyfact );
      _lineto_w( 360.00/(double)pixfact, 215.00/(double)piyfact );
    _setcolor(1);
      _lineto_w( 610.00/(double)pixfact, 215.00/(double)piyfact );
      _lineto_w( 610.00/(double)pixfact, 190.00/(double)piyfact );
    _setcolor(7);

    }

  settitletext();
  if(vc.numcolors > 2) _setcolor(0);
  _moveto( 90/pixfact, 194/piyfact);
  strcpy( title, "CURRENT DIVE No: ");
  itoa( (divenumber+1), titlednum, 10);
  strcat( title ,titlednum);
  _outgtext( title);

  _moveto( 400/pixfact, 194/piyfact);
  strcpy( title, "DIVE HISTORY");
  _outgtext( title);


  _setviewport( x1graph2/pixfact,y1graph2/piyfact, x2graph2/pixfact,y2graph2/piyfact );
  _setlinestyle( 0xffff);
  borderdraw();

  for(j=0; j<divenumber; j++) {
    for(i=0 ;i<=numberpoints[j]; i++) {
      if(depthmax < depthpoint[j][i]) depthmax = depthpoint[j][i];
      timetotal = timetotal + timepointc[j][i]+ timepointb[j][i]+ timepointa[j][i];
      /* printf("\ntimetotal=%f, depthmax=%f",timetotal,depthmax); */
    }
  }

  titleprofilegraph(xposgraph2, depthmax, timetotal );
  _setviewport( (x1graph2+5)/pixfact,y1graph2/piyfact, x2graph2/pixfact,(y2graph2-5)/piyfact );
  if( vc.numcolors > 2)_setcolor(14);
  _moveto_w( 0.00/(double)pixfact, 0.00/(double)piyfact);
  for(j=0; j<divenumber; j++) {
    for(i=0 ;i<=numberpoints[j]; i++) {
      if(bailoutpoint[j][i]) {
	if( vc.numcolors > 2) _setcolor(14);
      }
      else {
	if( vc.numcolors > 2) _setcolor(10);
      }
      _lineto_w( (timepointtotal+timepointa[j][i])*timetotalgraph, depthpoint[j][i]*depthmaxgraph);
      timepointtotal =	timepointtotal + timepointa[j][i];
      _lineto_w( (timepointtotal+timepointc[j][i]+timepointb[j][i])*timetotalgraph, depthpoint[j][i]*depthmaxgraph);
      timepointtotal =	timepointtotal + timepointb[j][i] +timepointc[j][i];
    }
  }

  tissuegraph();
       if(!strcmp(argvg,"bignose5")) {
	 while(!kbhit()); getch();
       }

  if( vc.numcolors > 2)_setcolor(0);
  _setviewport( x1graph1/pixfact,y1graph1/piyfact, x2graph1/pixfact,y2graph1/piyfact );
  borderdraw();

}

int licenseread(void)
{
int i;
char numtitle[50];
  if( (fp=fopen("diveplan.cop", "r+b")) == NULL ) {
    printf("Cannot open file");
    _setvideomode( _DEFAULTMODE);
    exit(1);
  }
  fread( licenseename, sizeof(char), (size_t) 100, fp);
  for(i=0; i<50; i++) licenseename[i] = licenseename[i] - 128; /* 0 to 50 */

  for(i=0; i<3; i++) {
    safetytitlednum[i] =licenseename[55+i];		      /* 55 to 59	*/
  }
  safetyfactor = ( (double)atoi(safetytitlednum) )/500.00 ;
  for(i=0; i<3; i++) {
    safetytitlednum[i] =licenseename[60+i];		      /* 60 to 64	*/
  }
  micro_mode = ( (double)atoi(safetytitlednum) )/100.00 ;
  if(micro_mode<=0.00 || micro_mode>1.00) micro_mode=1.00;

  for(i=0; i<7; i++) {					       /* 65 to 71 */
    safetytitlednum[i] =licenseename[65+i];
  }
  breathingrate = ( (double)atof(safetytitlednum) );
  if(!breathingrate) breathingrate=10.00;

  for(i=0; i<7; i++) {					       /* 72 to 71 */
    safetytitlednum[i] =licenseename[72+i];
  }
  atmospheric = ( (double)atof(safetytitlednum) );
  if(!atmospheric) atmospheric=1.00;

  for(i=0; i<10; i++) { 				       /* 80 to 89 */
    numtitle[i] =licenseename[80+i];
  }
  serialnumber = atol(numtitle);

  if(licenseename[90]=='f') {				       /* 90 */
    feetfactor = 3.28084;
    stopfactor = 3.048;
    strcpy ( form , "fsw");
    strcpy( formlong , "feet  ");
    psifactor = 14.7;
    strcpy ( porb , "psi");
    strcpy ( cuftorltrmin , "cuft/min");
    strcpy ( cuftorltr , "cuft  ");
    strcpy ( cuftorltrnos , "cuft ");
    maxbreathratenumber=2.00, minbreathratenumber=0.20;
    cuft_ltr_factor=28.316843;
  }

  if(licenseename[95]=='o') bubblesmode=0;		       /* 95 */

  if(licenseename[96]=='S') lasermode=1;		       /* 96 */
  if(licenseename[96]=='L') lasermode=2;

  if(licenseename[97]=='o') sixstopmode=1;		       /* 97 */
  if(licenseename[97]=='p') sixstopmode=2;		       /* 97 */

  if(TRADEGAS==2 && licenseename[98]=='g') {		       /* 98 */
      tradegasscreenmode=1;
  }

  fclose(fp);

  if( (fp=fopen("price.tot", "a+b")) == NULL ) {
    exit(3);
  }
  rewind(fp);
  fread( &airfillpricetotal, sizeof(double), (size_t) 1, fp);
  fread( &o2fillpricetotal, sizeof(double), (size_t) 1, fp);
  fread( &hefillpricetotal, sizeof(double), (size_t) 1, fp);
  fread( &airfillcosttotal, sizeof(double), (size_t) 1, fp);
  fread( &o2fillcosttotal, sizeof(double), (size_t) 1, fp);
  fread( &hefillcosttotal, sizeof(double), (size_t) 1, fp);
  fclose(fp);
}



int licensecheck(void)
{
int i;
  if( (fp=fopen("diveplan.cop", "r+b")) == NULL ) {
    printf("Cannot open file");
    _setvideomode( _DEFAULTMODE);
    exit(1);
  }

    _settextposition( 2,20);
    sprintf(stitle, "Thank you for purchasing PRO-DIVE PLANNER");
    _outtext( stitle) ;
    _settextposition( 3,17);
    sprintf(stitle, "Please enter your name (and company) to accept");
    _outtext( stitle) ;
    _settextposition( 4,16);
    sprintf(stitle, "the terms and conditions of use of this software");
    _outtext( stitle) ;
    _settextposition( 5,16);
    sprintf(stitle, "_________________________________________________");
    _outtext( stitle) ;
    _settextposition( 5,16);
    cnumbuf[0]=50;
    numbuf = cgetsa( cnumbuf );
    if(!*numbuf) {
      fclose(fp);
      exit (1);
    }
    strcpy( licenseename, numbuf );
    rewind(fp);
    for(i=0; i<50; i++) licenseename[i] = licenseename[i] + 128;
    fwrite( licenseename, sizeof(char), (size_t) 100, fp);
    fclose(fp);
    licenseread();
    return 1;
}

void licensefeetmodewrite(void)
{
int i;
  if( (fp=fopen("diveplan.cop", "r+b")) == NULL ) {
    printf("Cannot open file");
    _setvideomode( _DEFAULTMODE);
    exit(1);
  }
  fread( licenseename, sizeof(char), (size_t) 100, fp);

  itoa( ((int)(safetyfactor*500.00+0.49)), safetytitlednum, 10);
  for(i=0; i<3; i++) {
    licenseename[55+i] = safetytitlednum[i];	   /* 55 to 58 */
  }

  itoa( ((int)(micro_mode*100.00+0.49)), safetytitlednum, 10);
  for(i=0; i<3; i++) {
    licenseename[60+i] = safetytitlednum[i];	   /* 60 to 64 */
  }

  sprintf(safetytitlednum, "%5.2f", breathingrate);
  for(i=0; i<7; i++) {
    licenseename[65+i] = safetytitlednum[i];	   /* 65 to 71 */
  }

  sprintf(safetytitlednum, "%5.3f", atmospheric);
  for(i=0; i<7; i++) {
    licenseename[72+i] = safetytitlednum[i];	   /* 72 to 79 */
  }

  if(feetfactor > 1.00) licenseename[90] = 'f';
  else licenseename[90] = 'm';			   /* 90 */

  if(bubblesmode) licenseename[95]='b'; 	   /* 95 */
  else licenseename[95]='o';

  if(lasermode==1) licenseename[96]='S';	   /* 96 */
  if(lasermode==2) licenseename[96]='L';
  if(!lasermode) licenseename[96]='o';

  if(sixstopmode) licenseename[97]='n' + sixstopmode;	    /* 97 */
  else licenseename[97]='3';

  if(TRADEGAS==2 && tradegasscreenmode==1) licenseename[98]='g'; /* 98 */
  else licenseename[98]='o';

  rewind(fp);
  fwrite( licenseename, sizeof(char), (size_t) 100, fp);
  fclose(fp);
}

double stoptimetisscalc(int i, int fi, double stoptime0)
{
int j=0;
double tissuetry[3], timetry[3], ttissuetrymag;

  tissuetry[0] = tisscalc( i, timetry[0]=stoptime0+.50, fi);
  tissuetry[1] = tisscalc( i, timetry[1]=stoptime0+1.00, fi);
  tissuetrymag = fabs(tissuetry[1]);
  for(j=0; tissuetrymag > 0.00000001 && j < 1000 && !( tissuetry[1] == tissuetry[0] ); j++ ) {
    timetry[2] = timetry[1] - ( timetry[1] - timetry[0] ) * tissuetry[1]
					/ ( tissuetry[1] - tissuetry[0] );
    timetry[0] = timetry[1];
    timetry[1] = timetry[2];
    tissuetry[0] = tissuetry[1];
    pigttolstopminus1[i] = ( tolstoppressure ) / bcalc(i,toln2only,1) + acalc(i,toln2only,1);
    tissuetry[1] = tisscalc( i, timetry[1], fi);
    tissuetrymag = fabs(tissuetry[1]);
  }
  if( tissuetry[1] == tissuetry[0] ) return 29999.00;
  else return timetry[1];

}

double tisscalc(int i, double time, int fi)
{
  fractioncalcs(fi);

      algorithmcalc(i, time, fi);
      return (tissuetemphe[i] + tissuetemp[i] - pigttolstopminus1[i]);

}

void algorithmcalc(int i, double time, int fi)
{
      if(( nitrogenfractioncalc * absolutedepth - tissue[i] ) > 0)
	tissuetemp[i] = tissue[i] + ( nitrogenfractioncalc * absolutedepth - tissue[i] ) * (1.00 - exp((double)(-0.69316 * time / halftime[i] )));
      else if(fi)
	     tissuetemp[i] = tissue[i] + releaserate * ( nitrogenfractioncalc * absolutedepth - tissue[i] ) * (1.00 - exp((double)(-0.69316 * time / halftime[i] )));
	   else
	     tissuetemp[i] = tissue[i] + releaserate * 0.700 * ( nitrogenfractioncalc * absolutedepth - tissue[i] ) * (1.00 - exp((double)(-0.69316 * time / halftime[i] )));

      if(( heliumfractioncalc * absolutedepth - tissuehe[i] ) > 0)
	tissuetemphe[i] = tissuehe[i] + ( heliumfractioncalc * absolutedepth - tissuehe[i] ) * (1.00 - exp((double)(-0.69316 * time / halftimehe[i] )));
      else if(fi)
	     tissuetemphe[i] = tissuehe[i] + releaserate * ( heliumfractioncalc * absolutedepth - tissuehe[i] ) * (1.00 - exp((double)(-0.69316 * time / halftimehe[i] )));
	   else
	     tissuetemphe[i] = tissuehe[i] + releaserate * 0.700 * ( heliumfractioncalc * absolutedepth - tissuehe[i] ) * (1.00 - exp((double)(-0.69316 * time / halftimehe[i] )));
}

void ppo2print(double pp)
{
double ppidoub;
  ppf = (int)(100.00 * (modf(pp+0.005, &ppidoub) ) ) ;
  ppi = (int)ppidoub;
}

void setsafetyfactor(void)
{
double sftemp;
  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SET SAFETY");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("  FACTOR");
  if( vc.numcolors > 2) _setcolor(7);
  helpscreen(19);

  do {
      _settextposition( 14,25);
      sprintf(stitle, "Current Safety factor %d%%  ", (int)(safetyfactor*500.00) );
      _outtext( stitle) ;
      _settextposition( 15,25);
      sprintf(stitle, "Enter Safety factor(0 to 50%%)= __%%  ");
      _outtext( stitle) ;
      _settextposition( 15,56);
      cnumbuf[0]=3;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) break;
      sftemp = ( (double)atoi( numbuf ) ) / 500.00 ;
  } while( (sftemp > .10) || !*numbuf);
  if(*numbuf) safetyfactor = sftemp;

  do {
      _settextposition( 14,25);
      sprintf(stitle, "Current Micro Bubble factor %d%%  ", (int)(micro_mode*100.00) );
      _outtext( stitle) ;
      _settextposition( 15,25);
      sprintf(stitle, "Enter Micro Bubble factor(0 to 100%%)= ___%%  ");
      _outtext( stitle) ;
      _settextposition( 15, 63);
      cnumbuf[0]=4;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atoi( numbuf ) ) / 100.00 ;
  } while( (sftemp > 1.0) || !*numbuf);
  micro_mode = sftemp;
  //safetyfactor+micro_mode*18.00/500.00;

}

void setatmospheric(void)
{
double sftemp;
  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SET ATMOSPHERIC");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("   PRESSURE");
  if( vc.numcolors > 2) _setcolor(7);
  helpscreen(20);

  do {
      _settextposition( 14,25);
      sprintf(stitle, "Current Atmospheric pressure= %d mBar  ", (int)(atmospheric*1000.00+0.49) );
      _outtext( stitle) ;
      _settextposition( 15,25);
      sprintf(stitle, "Enter Atmospheric pressure= ____ mBar  ");
      _outtext( stitle) ;
      _settextposition( 15,53);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atoi( numbuf ) ) / 1000.00 ;
  } while( (sftemp > 1.50) || !*numbuf);
  atmospheric = sftemp;

}

void setbreathingrate(void)
{
double sftemp;
  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SET BREATHING");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("       RATE");
  if( vc.numcolors > 2) _setcolor(7);
  helpscreen(23);

  do {
      _settextposition( 14,25);
      sprintf(stitle, "Current Breathing rate= %3.2f%s",feetfactor==1.00 ? breathingrate : breathingrate/cuft_ltr_factor, cuftorltrmin);
      _outtext( stitle) ;
      _settextposition( 15,25);
      sprintf(stitle, "Enter Breathing rate= ____%s  ",cuftorltrmin);
      _outtext( stitle) ;
      _settextposition( 15,47);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atof( numbuf ) );
      if( (sftemp > maxbreathratenumber) ) helpscreen(24);
      if( (sftemp < minbreathratenumber) ) helpscreen(25);
  } while( (sftemp < minbreathratenumber) || (sftemp > maxbreathratenumber) || !*numbuf);
  if(feetfactor!=1.00) breathingrate = sftemp * cuft_ltr_factor;
  else breathingrate = sftemp;

}


void setmissionstart(void)
{
int sftemp, sftemp2;
  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SET MISSION START");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("  TIME AND DAY");
  if( vc.numcolors > 2) _setcolor(7);

  do {
      _settextposition( 14,25);
      sprintf(stitle, "Current Mission start time= %02d:%02d  ", missionstart[1], missionstart[2] );
      _outtext( stitle) ;
      _settextposition( 15,25);
      sprintf(stitle, "Enter Mission start time __:__     ");
      _outtext( stitle) ;
      helpscreen(21);
      _settextposition( 15,50);
      cnumbuf[0]=3;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = atoi( numbuf )  ;
      helpscreen(22);
      _settextposition( 15,53);
      cnumbuf[0]=3;
      numbuf = cgetsn( cnumbuf, "", "" );
      sftemp2 = atoi( numbuf ) ;
  } while( (sftemp > 23) || (sftemp2 > 59) || !*numbuf);
  missionstart[0] = 1;
  missionstart[1] = sftemp;
  missionstart[2] = sftemp2;
}

int getdiscsaveddive(void)
{
int fok;
struct find_t find;
unsigned char title[100];

  if(!strcmp(argvgn[0], "pass")) {
    if(argvgn[1][0]) {
      strcpy(title, argvgn[1]);
      strcpy(divetitl, argvgn[1]);
      strcat(title,".div");
      if( !(fp=fopen(title, "rb")) ) {
	printf("Cannot open source file");
	fok =1;
	return 0;
	/*exit(1);*/
      }
    }
  }
  else {
    drawbackground();
    _moveto( 100/pixfact, 11/piyfact);
    _outgtext("PRINT DIVE");
    _moveto( 100/pixfact, 26/piyfact);
    _outgtext("");
    if( vc.numcolors > 2) _setcolor(7);
    divelist();
    helpscreen(18);
    do {
      fok = 0;
      _settextposition( 4,30);
      sprintf(stitle, "Enter Dive name: ________");
      _outtext( stitle) ;
      _settextposition( 5,30);
      sprintf(stitle, "(return for current dive)");
      _outtext( stitle) ;
      _settextposition( 4,47);
      cnumbuf[0]=9;
      numbuf = cgetsa( cnumbuf );
      if(!*numbuf) return 0;
      strcpy(title, numbuf);
      strcpy(divetitl, numbuf);
      strcat(title,".div");
      if( !(fp=fopen(title, "rb")) ) {
	printf("Cannot open source file");
	fok =1;
	/*exit(1);*/
      }
    } while( fok );
    if( !_dos_findfirst( title, 0xffff, &find ) ) {
	fileinfo( &find );
	/*
	_settextposition( j,5);
	sprintf(stitle,	"  %-8s",filebuf);
	_outtext( stitle) ;
	_settextposition( j+1,5);
	sprintf(stitle,	"  %-8s",datebuf);
	_outtext( stitle) ;
	*/
    }
  }

  fread( &divenumber, sizeof(int), (size_t) 1, fp);
  fread( numberpoints, sizeof(int), (size_t) 10, fp);
  fread( ppo2exptime_14, sizeof(double), (size_t) 10, fp);
  fread( ppo2exptime_15, sizeof(double), (size_t) 10, fp);
  fread( ppo2exptime_16, sizeof(double), (size_t) 10, fp);
  fread( ppo2exptime_16plus, sizeof(double), (size_t) 10, fp);
  fread( &safetyfactordive, sizeof(int), (size_t) 1, fp);
  fread( &atmosphericdive, sizeof(int), (size_t) 1, fp);
  fread( &air, sizeof(int), (size_t) 1, fp);
  fread( &n2, sizeof(int), (size_t) 1, fp);
  fread( &he, sizeof(int), (size_t) 1, fp);
  fread( &ppo2, sizeof(int), (size_t) 1, fp);
  fread( divestart, sizeof(int), (size_t) 30, fp);
  fread( divefinish, sizeof(int), (size_t) 30, fp);
  fread( flytol, sizeof(int), (size_t) 3, fp);
  fread( ppo2cns, sizeof(double), (size_t) 10, fp);
  fread( totaltimetosurface, sizeof(long), (size_t) 10, fp);
  fread( ppo2cnsmax, sizeof(double), (size_t) 10, fp);
  fread( &sixstopmodedive, sizeof(int), (size_t) 1, fp);
  fread( maxppo2, sizeof(double), (size_t) 10, fp);
  fread( diveotu, sizeof(double), (size_t) 10, fp);
  fread( &missionotu, sizeof(double), (size_t) 1, fp);
  fread( gasmix, sizeof(double), (size_t) 20*NUMGASMIX, fp);
  fread( gasmixbartime, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( gasreservefraction, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( filldive, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( fillres, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( filltotal, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( cylindersize, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( maxcylinderpressure, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( freecylindersize, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fread( &breathingratedive, sizeof(double), (size_t) 1, fp);
  fread( &toln2only, sizeof(int), (size_t) 1, fp);
  fread( timetofirststop, sizeof(long), (size_t) 10, fp);
  fread( missiontotal, sizeof(int), (size_t) 3, fp);

  fread( tissue, sizeof(double), (size_t) 16, fp);
  fread( tissuehe, sizeof(double), (size_t) 16, fp);

  fread( depthpoint, sizeof(double), (size_t) 1100, fp);
  fread( timepointa, sizeof(double), (size_t) 1100, fp);
  fread( timepointb, sizeof(double), (size_t) 1100, fp);
  fread( nitrogenpoint, sizeof(double), (size_t) 1100, fp);
  fread( heliumpoint, sizeof(double), (size_t) 1100, fp);
  fread( ppo2point, sizeof(double), (size_t) 1100, fp);
  fread( hemixpoint, sizeof(double), (size_t) 1100, fp);
  fread( n2mixpoint, sizeof(double), (size_t) 1100, fp);
  fread( bailoutpoint, sizeof(int), (size_t) 1100, fp);
  fread( timepointc, sizeof(double), (size_t) 1100, fp);
  fread( stoptimeplus, sizeof(double), (size_t) 1100, fp);
  fread( &micro_factordive, sizeof(int), (size_t) 1, fp);


  fclose(fp);
  atmospheric=(float)atmosphericdive/1000.00;
  strcpy(title, divetitl);
  strcat(title, ".mix");
  getdisc_gasmixdata(title);
  return 1;
}


void savediveondisc2(void)
{
int fok;
struct find_t find;
unsigned char title[100];

  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SAVE DIVE");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext(" TO DISC");
  if( vc.numcolors > 2) _setcolor(7);
  divelist();
  helpscreen(17);
  do {
      fok = 0;
      _settextposition( 4,30);
      sprintf(stitle, "Enter Dive name: ________");
      _outtext( stitle) ;
      _settextposition( 4,47);
      cnumbuf[0]=9;
      numbuf = cgetsa( cnumbuf );
      if(!*numbuf) return;
      strcpy(title, numbuf);
      strcpy(divetitl, numbuf);
      strcat(title,".div");
      if( (fp=fopen(title, "w+b")) == NULL ) {
	printf("Cannot open source file");
	fok =1;

	/*exit(1);*/
      }
  } while( fok );
  if( !_dos_findfirst( title, 0xffff, &find ) ) {
	fileinfo( &find );
	/*
	_settextposition( j,5);
	sprintf(stitle,	"  %-8s",filebuf);
	_outtext( stitle) ;
	_settextposition( j+1,5);
	sprintf(stitle,	"  %-8s",datebuf);
	_outtext( stitle) ;
	*/
  }

  fwrite( &divenumber, sizeof(int), (size_t) 1, fp);
  fwrite( numberpoints, sizeof(int), (size_t) 10, fp);
  fwrite( ppo2exptime_14, sizeof(double), (size_t) 10, fp);
  fwrite( ppo2exptime_15, sizeof(double), (size_t) 10, fp);
  fwrite( ppo2exptime_16, sizeof(double), (size_t) 10, fp);
  fwrite( ppo2exptime_16plus, sizeof(double), (size_t) 10, fp);
  fwrite( &safetyfactordive, sizeof(int), (size_t) 1, fp);
  fwrite( &atmosphericdive, sizeof(int), (size_t) 1, fp);
  fwrite( &air, sizeof(int), (size_t) 1, fp);
  fwrite( &n2, sizeof(int), (size_t) 1, fp);
  fwrite( &he, sizeof(int), (size_t) 1, fp);
  fwrite( &ppo2, sizeof(int), (size_t) 1, fp);
  fwrite( divestart, sizeof(int), (size_t) 30, fp);
  fwrite( divefinish, sizeof(int), (size_t) 30, fp);
  fwrite( flytol, sizeof(int), (size_t) 3, fp);
  fwrite( ppo2cns, sizeof(double), (size_t) 10, fp);
  fwrite( totaltimetosurface, sizeof(long), (size_t) 10, fp);
  fwrite( ppo2cnsmax, sizeof(double), (size_t) 10, fp);
  fwrite( &sixstopmodedive, sizeof(int), (size_t) 1, fp);
  fwrite( maxppo2, sizeof(double), (size_t) 10, fp);
  fwrite( diveotu, sizeof(double), (size_t) 10, fp);
  fwrite( &missionotu, sizeof(double), (size_t) 1, fp);
  fwrite( gasmix, sizeof(double), (size_t) 20*NUMGASMIX, fp);
  fwrite( gasmixbartime, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( gasreservefraction, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( filldive, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( fillres, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( filltotal, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( cylindersize, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( maxcylinderpressure, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( freecylindersize, sizeof(double), (size_t) 10*NUMGASMIX, fp);
  fwrite( &breathingratedive, sizeof(double), (size_t) 1, fp);
  fwrite( &toln2only, sizeof(int), (size_t) 1, fp);
  fwrite( timetofirststop, sizeof(long), (size_t) 10, fp);
  fwrite( missiontotal, sizeof(int), (size_t) 3, fp);

  fwrite( tissue, sizeof(double), (size_t) 16, fp);
  fwrite( tissuehe, sizeof(double), (size_t) 16, fp);

  fwrite( depthpoint, sizeof(double), (size_t) 1100, fp);
  fwrite( timepointa, sizeof(double), (size_t) 1100, fp);
  fwrite( timepointb, sizeof(double), (size_t) 1100, fp);
  fwrite( nitrogenpoint, sizeof(double), (size_t) 1100, fp);
  fwrite( heliumpoint, sizeof(double), (size_t) 1100, fp);
  fwrite( ppo2point, sizeof(double), (size_t) 1100, fp);
  fwrite( hemixpoint, sizeof(double), (size_t) 1100, fp);
  fwrite( n2mixpoint, sizeof(double), (size_t) 1100, fp);
  fwrite( bailoutpoint, sizeof(int), (size_t) 1100, fp);
  fwrite( timepointc, sizeof(double), (size_t) 1100, fp);
  fwrite( stoptimeplus, sizeof(double), (size_t) 1100, fp);
  fwrite( &micro_factordive, sizeof(int), (size_t) 1, fp);

  fclose(fp);
  strcpy(title, divetitl);
  strcat(title, ".mix");
  putdisc_gasmixdata(title);

}

void ppo2exposuretime(void)
{
int i;

  ppo2_now = absolutedepthpure * ( 1.00 - nitrogenfraction - heliumfraction);
  if(ppo2_now >= 1.70) ppo2exptime_16plus[divenumber] = ppo2exptime_16plus[divenumber] + exposuretime;
  else if(ppo2_now >= 1.60) ppo2exptime_16[divenumber] = ppo2exptime_16[divenumber] + exposuretime;
       else if(ppo2_now >= 1.50) ppo2exptime_15[divenumber] = ppo2exptime_15[divenumber] + exposuretime;
	    else if(ppo2_now >= 1.40) ppo2exptime_14[divenumber] = ppo2exptime_14[divenumber] + exposuretime;

  for(i=0; (ppo2_now>cnslooktable[i][0]) && i<55; i++);
  ppo2cnscurrent = ppo2cnscurrent + ( cnslooktable[i][1] * exposuretime);
  if( (i>54) || ppo2cnscurrent > 9999.00 ) ppo2cnscurrent = 9999.00;
  if(ppo2cnscurrent<0.00) ppo2cnscurrent = 0.00;
  if( ppo2cnsmax[divenumber] < ppo2cnscurrent ) ppo2cnsmax[divenumber] = ppo2cnscurrent;
}

void tissuegraph(void)
{
int i, j, ptisstolatmospercenty[16];
double ptisstolatmospercent[16], heptisstolatmospercent[16];
unsigned char title[100];
long imsize;

  if( piyfactdec) return; /*return if y pixels <220 */

  _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );
  setaxistext();
  if( vc.numcolors > 2) _setcolor(0);
  _moveto( 260/pixfact, 389/piyfact);
  strcpy( title, "TISSUE STATUS");
  _outgtext( title);
  setmicroaxistext();
  if( vc.numcolors > 2) {
    _setcolor(15);
    _moveto_w( 400.00/(double)pixfact, 387.00/(double)pixfact);
      _lineto_w( 250.00/(double)pixfact, 387.00/(double)piyfact );
      _lineto_w( 250.00/(double)pixfact, 404.00/(double)piyfact );
    _setcolor(1);
      _lineto_w( 400.00/(double)pixfact, 404.00/(double)piyfact );
      _lineto_w( 400.00/(double)pixfact, 387.00/(double)piyfact );

    _setcolor(8);
    x=(x1graph1+6)/pixfact; y=(480/piyfact)-75/piyfact;
    _rectangle( _GFILLINTERIOR, x2graph2/pixfact, (480/piyfact)-25/piyfact, x, y );
/*    imsize = _imagesize( x, y, x2graph2/pixfact, (480/piyfact)-25/piyfact );
    buffert = halloc( imsize, 1 );
    _getimage( x, y, x2graph2/pixfact, (480/piyfact)-25/piyfact , buffert );
*/
    _setcolor(7);
  }
  for(i=0, j=34; i<16; i++, j=j+34) {
    ptisstolatmospercent[i]  =	100.00 * (tissue[i] + tissuehe[i] - (0.79*atmospheric)) /
						       ( (atmospheric / bcalc(i,toln2only,0)) + acalc(i,toln2only,0) - (0.79*atmospheric) ) ;
    if(ptisstolatmospercent[i] < 0.00) {
      ptisstolatmospercent[i] = 0.00;
    }
    ptisstolatmospercenty[i] =	(int)(0.50 * ptisstolatmospercent[i])/piyfact ;
    if(ptisstolatmospercent[i] >= 100.00) {
      ptisstolatmospercenty[i] = 50/piyfact;
    }
    if(!strcmp(argvg,"bignose5")) {
      printf("\nptiss=%g, ptissh=%g",tissue[i],tissuehe[i]);
    }

    if( vc.numcolors > 2) _setcolor(12);
    x=(x1graph1+6+j)/pixfact; y=(480/piyfact)-25;
    _rectangle( _GFILLINTERIOR, (x1graph1+6+j-32)/pixfact, (480/piyfact)-25-ptisstolatmospercenty[i], x, y );
    heptisstolatmospercent[i]  =	100.00 * (tissue[i]+tissuehe[i] - (0.79*atmospheric)) /
						       ( (atmospheric / bcalc(i,toln2only,0)) + acalc(i,toln2only,0) - (0.79*atmospheric) ) ;
    if(heptisstolatmospercent[i] < 0.00) {
      heptisstolatmospercent[i] = 0.00;
    }
    ptisstolatmospercenty[i] =	(int)(0.50 * heptisstolatmospercent[i])/piyfact ;
    if(ptisstolatmospercent[i] >= 100.00) {
      ptisstolatmospercenty[i] = 50/piyfact;
    }
    ptisstolatmospercenty[i] = (ptisstolatmospercenty[i]*(int)(100.00*tissuehe[i]/(tissue[i]+tissuehe[i])))/100;
    if( vc.numcolors > 2) _setcolor(11);
    x=(x1graph1+6+j)/pixfact; y=(480/piyfact)-25;
    _rectangle( _GFILLINTERIOR, (x1graph1+6+j-32)/pixfact, (480/piyfact)-25-ptisstolatmospercenty[i], x, y );

    if( vc.numcolors > 2) _setcolor(15);
    itoa( (int)ptisstolatmospercent[i], title, 10);
    strcat( title , "%");
    _moveto( (x1graph1+6+j-30)/pixfact, 405/piyfact);
    _outgtext( title);

    _moveto( (x1graph1+6+2)/pixfact, 425/piyfact);
    _outgtext("FAST");
    _moveto( (x1graph1+6+504)/pixfact, 425/piyfact);
    _outgtext("SLOW");


  }

}

void divelist(void)
{
    int i=1, j=6;
    struct find_t find;
    long size;

    /* Find first matching file, then find additional matches. */
    _settextposition( j,5);
    sprintf(stitle, "  Dive list:");
    _outtext( stitle) ;
    j++;
    if( !_dos_findfirst( "*.div", 0xffff, &find ) )
    {
	size = fileinfo( &find );
	_settextposition( j,5);
	sprintf(stitle,	"  %-8s",filebuf);
	_outtext( stitle) ;
	_settextposition( j+1,5);
	sprintf(stitle,	"  %-8s",datebuf);
	_outtext( stitle) ;
    }
    while( !_dos_findnext( &find ) ) {
	if(j>27) {
	  j=6;
	  _settextposition( j,3+i*11);
	  _outtext( "press any key for next page") ;
	  getch();
	  _settextposition( j,3+i*11);
	  _outtext( "                           ") ;
	  j++;
	}
	size += fileinfo( &find );
	_settextposition( j,5+i*10);
	sprintf(stitle,	"  %-8s",filebuf);
	_outtext( stitle) ;
	_settextposition( j+1,5+i*10);
	sprintf(stitle,	"  %-8s",datebuf);
	_outtext( stitle) ;
	i++;
	if(i>6) {
	  i=0;
	  j++;
	  j++;
	}
    }
}

/* Displays information about a file. */
long fileinfo( struct find_t *pfind )
{
    int i;

    datestr( pfind->wr_date, datebuf );
    timestr( pfind->wr_time, timebuf );

    strcpy(filebuf, pfind->name);
    for(i=0; i<9; i++){
      if(filebuf[i] == '.') {
	filebuf[i] = '\0';
	break;
      }
    }
    return pfind->size;
}

/* Takes unsigned time in the format:		    fedcba9876543210
 * s=2 sec incr, m=0-59, h=23			    hhhhhmmmmmmsssss
 * Changes to a 9-byte string (ignore seconds):     hh:mm ?m
 */
char *timestr( unsigned t, char *buf )
{
    int h = (t >> 11) & 0x1f, m = (t >> 5) & 0x3f;

    sprintf( buf, "%2.2d:%02.2d %cm", h % 12, m,  h > 11 ? 'p' : 'a' );
    return buf;
}

/* Takes unsigned date in the format:		    fedcba9876543210
 * d=1-31, m=1-12, y=0-119 (1980-2099)		    yyyyyyymmmmddddd
 * Changes to a 9-byte string:			    mm/dd/yy
 */
char *datestr( unsigned d, char *buf )
{
    if(feetfactor==1.00)
      sprintf( buf, "%2.2d/%02.2d/%02.2d",
	     d & 0x1f, (d >> 5) & 0x0f, (d >> 9) + 80 );
    else
      sprintf( buf, "%2.2d/%02.2d/%02.2d",
	     (d >> 5) & 0x0f, d & 0x1f, (d >> 9) + 80 );
    return buf;
}



void missiontotalupdate( double timeinc )
{
long mth, mtm, mtd;
  mtm = (long)missiontotal[2] + (long)(exposuretime + timeinc);
  mth = mtm/60;
  mtm = mtm%60;
  mth = mth + (long)missiontotal[1];
  mtd = mth/24 + (long)missiontotal[0];
  mth = mth%24;
  missiontotal[0] = (int)mtd;
  missiontotal[1] = (int)mth;
  missiontotal[2] = (int)mtm;
}

void flytolupdate(void)
{
long mth, mtm, mtd;
  mtm = (long)missiontotal[2] + (long)flytolmins;
  mth = mtm/60;
  mtm = mtm%60;
  mth = mth + (long)missiontotal[1];
  mtd = mth/24 + (long)missiontotal[0];
  mth = mth%24;
  flytol[0] = (int)mtd;
  flytol[1] = (int)mth;
  flytol[2] = (int)mtm;
}

void bubbles( void)
{
long imsize;
    if(!bubblesmode) return;
    /* Measure the image to be drawn and allocate memory for it. */
    imsize = _imagesize( -3, -3, +3, +3 );
    buffer = halloc( imsize, 1 );
    if( buffer == NULL )
	exit( 1 );

    if( !(fp2=fopen("image2.bit", "rb")) ) {
      printf("Cannot open source file");
      getch();
      hfree( buffer );
      exit(1);
    }
    fread( buffer, sizeof(char), (size_t) imsize, fp);
    fclose(fp2);

    bubblex=600; bubbley=480;
    bubbley=bubbley+10;
    pubub(); /* 490 */
    bubbley=bubbley-240;
    pubub(); /* 250 */
    bubbley=480;
    while( 1 )
      {
	  pubub();	 /* 480 */    /* 10 */
	  bubbley=bubbley+10;
	  pubub();	 /* 490 */    /* 20 */
	  bubbley=bubbley-250;
	  pubub();	 /* 240 */    /* 250 */
	  bubbley=bubbley+10;
	  pubub();	 /* 250 */    /* 260 */
	  bubbley=bubbley+220;
	  bubbley= !bubbley ? 480 : bubbley ;
	  if(kbhit()) {
	    hfree( buffer );
	    return;
	  }
    }
}

void pubub(void)
{
int j;
short xrand[20] = {
  245,15,285,590,400,570,310,385,500,420,120,450,100,540,140,80,220,170,40,470
  };
short yrand[20] = {
  190,410,375,20,320,120,210,350,410,250,90,390,40,290,150,50,35,280,350,20
  };

  for(j=0;j<15;j++) {
    _putimage( bubblex-xrand[j] , (bubbley-yrand[j])>=0 ? bubbley-yrand[j] :	(480+bubbley-yrand[j])>=0 ? 480+bubbley-yrand[j] : 960+bubbley-yrand[j]	, buffer, action[0] );
    _putimage( bubblex-xrand[j] , (bubbley-yrand[j]-1)>=0 ? bubbley-yrand[j]-1 :	(480+bubbley-yrand[j]-1)>=0 ? 480+bubbley-yrand[j]-1 : 960+bubbley-yrand[j]-1	, buffer, action[0] );
  }
}


void abhalffile(void)
{

  if(!strcmp(argvg,"w")) {
    if( (fp=fopen("abhalf.con", "wb")) == NULL ) {
      printf("Cannot open file");
      return;
    }
    fwrite( a, sizeof(double), (size_t) 16, fp);
    fwrite( b, sizeof(double), (size_t) 16, fp);
    fwrite( halftime, sizeof(double), (size_t) 16, fp);
    fclose(fp);
    return;
  }

  if(!strcmp(argvg,"r")) {
    if( (fp=fopen("abhalf.con", "rb")) == NULL ) {
      printf("Cannot open file");
      return;
    }
    fread( a, sizeof(double), (size_t) 16, fp);
    fread( b, sizeof(double), (size_t) 16, fp);
    fread( halftime, sizeof(double), (size_t) 16, fp);
    fclose(fp);
    return;
  }

}

void drawbackground(void)
{

  _clearscreen( _GCLEARSCREEN);
  if( vc.numcolors > 2) {
    _setlinestyle( 0xffff);
    _setcolor(15);
    _moveto_w( (double)vc.numxpixels-10.00/(double)pixfact, 10.00/(double)pixfact);
      _lineto_w( 10.00/(double)pixfact, 10.00/(double)piyfact );
      _lineto_w( 10.00/(double)pixfact, (double)(480/piyfact)-10.00/(double)piyfact );
    _setcolor(7);
      _lineto_w( (double)vc.numxpixels-10.00/(double)pixfact, (double)(480/piyfact)-10.00/(double)piyfact );
      _lineto_w( (double)vc.numxpixels-10.00/(double)pixfact, 10.00/(double)piyfact );

    _setcolor(1);
    _moveto_w( (double)vc.numxpixels-11.00/(double)pixfact, 45.00/(double)pixfact);
      _lineto_w( 11.00/(double)pixfact, 45.00/(double)piyfact );

    _setcolor(7);
     x=11/pixfact; y=11/piyfact;
     _rectangle( _GFILLINTERIOR, vc.numxpixels-11/pixfact, 44/piyfact, x, y );
    _setcolor(0);
  }

}

void backgroundtoggle( void)
{
  if(vc.numcolors > 2) {
    x=(x1graph2+6)/pixfact; y=y1graph2/piyfact;
    _putimage( x, y, bufferg, action[0] ); /*, x+(x2graph2-x1graph2-6)/pixfact, y+(y2graph2-y1graph2-6)/piyfact*/
    x=(x1graph1+6)/pixfact; y=y1graph1/piyfact;
    _putimage( x, y, bufferg, action[0] );
    x=(x1graph1+6)/pixfact; y=(480/piyfact)-75/piyfact;
    _putimage( x, y, buffert, action[0] ); /*, x2graph2/pixfact, (480/piyfact)-25/piyfact*/
  }
}

void open_output( void)
{
  if((fprn=fopen( "PRN", "w"))==NULL) {
    _outtext( "Unable to open printer file.");
    exit(1);
  }
}

void activate_graphic_mode( void)
{
  int f;
  short unsigned code[4] = { 27, 76, 64, 3};

  for( f=0; f<4; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}

void activate_lasergraphic_mode( void)
{
  int f;
  short unsigned code[7] = { 27, '*', 't', '1', '0', '0', 'R'};

  for( f=0; f<7; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}

void activate_lasergraphic_mode75( void)
{
  int f;
  short unsigned code[7] = { 27, '*', 't', '7', '5', 'R'};

  for( f=0; f<7; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}

void deactivate_lasergraphic_mode( void)
{
  int f;
  short unsigned code[4] = { 27, '*', 'r', 'B'};

  for( f=0; f<4; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}

void laserleftmargin( void)
{
  int f;
  short unsigned code[5] = { 27, '*', 'r', '1', 'A'};

  for( f=0; f<5; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}

void lasernumberofbytes( int bytenum)
{
  int f;
  unsigned char titlednum[10];
  short unsigned code[6] = { 27, '*', 'b', '8', '0', 'W'};

  for( f=0; f<3; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);

  itoa(bytenum,titlednum,10);
  for( f=0; titlednum[f]; f++)
    fwrite( &titlednum[f], sizeof( char), (size_t) 1, fprn);

  fwrite( &code[5], sizeof( char), (size_t) 1, fprn);
}

void laserposition( int xl, int yl)
{
  int f;
  unsigned char titlednum[10];
  short unsigned xcode[7] = { 27, '*', 'p', '1', '5', '0', 'X'};
  short unsigned ycode[5] = { 27, '*', 'p', '0', 'Y'};

  for( f=0; f<7; f++)
    fwrite( &xcode[f], sizeof( char), (size_t) 1, fprn);

  for( f=0; f<3; f++)
    fwrite( &ycode[f], sizeof( char), (size_t) 1, fprn);

  itoa(yl,titlednum,10);
  for( f=0; titlednum[f]; f++)
    fwrite( &titlednum[f], sizeof( char), (size_t) 1, fprn);

  fwrite( &ycode[4], sizeof( char), (size_t) 1, fprn);
}


void send_linefeed( void)
{
  int f;
  short unsigned code[4] = { 27, 51, 24, 10};

  for( f=0; f<4; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}


void restore_linefeed( void)
{
  int f;
  short unsigned code[5] = { 27, 65, 12, 27, 50};

  for( f=0; f<5; f++)
    fwrite( &code[f], sizeof( char), (size_t) 1, fprn);
}

void delay1sec(void)
{
  clock_t  cstart, cend;    /* For clock	      */
  time_t   tstart, tend;    /* For difftime	      */

    /* time( &tstart );	*/ /* Use time and difftime for timing to seconds   */
    /* time( &tend );	*/
    /* while ( !difftime( tend, tstart ) ) time( &tend );
    */
    cstart = clock();	 /* Use clock for timing to hundredths of seconds */
    cend = clock();
    while ( (((float)(cend - cstart) ) / CLOCKS_PER_SEC ) < 1.00 ) cend = clock();
}
