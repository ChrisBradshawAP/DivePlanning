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

extern FILE _huge *fp2;

extern short action[5];
extern char *descrip[5];
extern char _huge *buffer, *bufferg, *buffert;
extern unsigned char divetitl[100];
extern char timebuf[10], datebuf[10], tradedatebuf[128], filebuf[20];

extern unsigned char licenseename[100], argvg[20], divername[20][20];

extern unsigned char *str[18];
extern unsigned char *str2[20];
extern double cnslooktable[55][2];
extern	 short pixfact, piyfact, piyfactdec;
  /*unsigned char menunumber;*/
extern   unsigned char options[8];

extern   unsigned char safetytitlednum[10];
extern	 unsigned char list[20];
extern   char fondir[_MAX_PATH];
extern   struct videoconfig vc;
extern   struct _fontinfo fi;
extern   short x, y, f;
extern   long prev_bk;
extern   int xint, yint;

extern char cnumbuf[MAXSTR];
extern char tmxbuf[MAXSTR];
extern char  *numbuf;
extern char form[4], formlong[10], porb[4], cuftorltrmin[9], cuftorltr[9], cuftorltrnos[9], gascurrency[9];
extern unsigned char stitle[80];
extern short bubblex, bubbley;

extern int air, n2, he, ppo2, ppi, ppf, number_gasses;
extern int x1graph1, y1graph1, x2graph1, y2graph1, x1graph2, y1graph2, x2graph2, y2graph2, xposgraph1, xposgraph2;
extern int tissueno, depthno, deepeststop, numberstops, divenumber, numberpoints[10], divestart[10][3], divefinish[10][3], flytol[3];
extern long serialnumber, totaltimetosurface[10], timetofirststop[10];
extern double ppo2exptime_14[10], ppo2exptime_15[10], ppo2exptime_16[10], ppo2exptime_16plus[10], ppo2_now, ppo2fractionlast, ppo2cns[10], ppo2cnsmax[10], ppo2cnscurrent;
extern double tissue[16], a[16], b[16], halftime[16], tissuetemp[16], pambtoltiss[16], stoptimetissue[16], stoptime[110];
extern double tissuehe[16], ahe[16], bhe[16], halftimehe[16], tissuetemphe[16], pambtoltisshe[16], stoptimetissuehe[16];
extern double absolutedepth, absolutedepthpure, depth, atmospheric, exposuretime, flytolmins, pambtolmin, pambtolmindepth, pigttolstopminus1[16], nitrogenfraction, nitrogenfracdec[110], heliumfraction, heliumfracdec[110];
extern double ppo2fraction, ppo2fracdec[110], surftime, depthlast, currentstoppressure;
extern double nitrogenfractiondepth[10], exposuretimedepth[10], depthdepth[10], heliumfractiondepth[10], ppo2fractiondepth[10];
extern double hemixpointfractiondepth[10], n2mixpointfractiondepth[10];
extern double depthpoint[10][110], timepointa[10][110], timepointb[10][110], timepointc[10][110], nitrogenpoint[10][110], heliumpoint[10][110], ppo2point[10][110], totaltimepoint[10][110], hemixpoint[10][100], n2mixpoint[10][100];
extern double hemixpointfracdec[110], n2mixpointfracdec[110];
extern double depthmaxgraph, timetotalgraph, dailyotu[50], diveotu[10], missionotu, maxppo2[10], safetyfactor;
extern double feetfactor, stopfactor, cuft_ltr_factor, psifactor, maxdepthalarm, nitrogenfractioncalc, heliumfractioncalc;
/*double depthinc, decdepthinc, releaserate, ptolinc; Tom release */
extern double depthinc, decdepthinc, releaserate, ptolinc;
extern double nitrogenfractionlast, heliumfractionlast, fractionmax;
extern double gasmix[11][NUMGASMIX][3], gasmixbartime[10][NUMGASMIX], gasreservefraction[10][NUMGASMIX];
extern double gasmixtable[NUMGASMIX][3];
extern char gasstatus[NUMGASMIX], gasused[NUMGASMIX], set_gastable;
extern double filldive[10][NUMGASMIX], fillres[10][NUMGASMIX], filltotal[10][NUMGASMIX], cylindersize[10][NUMGASMIX], maxcylinderpressure[10][NUMGASMIX], freecylindersize[10][NUMGASMIX];
extern double breathingrate, breathingratedive, maxbreathratenumber, minbreathratenumber;
extern int  atmosphericdive, safetyfactordive, menunumber, missionstart[3], missiontotal[3], bubblesmode, sixstopmode, lasermode, sixstopmodedive;
extern short unsigned sprite[720];
extern double cdiff;
extern double atotal, btotal, nfrac, hefrac, tolstoppressure, currentfillpressure;
extern int toln2only;
extern double heprice, o2price, airprice, hecost, o2cost, aircost;
extern double tradecylindersize, tradefreecylindersize, trademaxcylinderpressure;
extern double airfillltr, hefillltr, o2fillltr, airfillprice, hefillprice, o2fillprice;
extern double airfillpricetotal, hefillpricetotal, o2fillpricetotal;
extern double airfillcosttotal, hefillcosttotal, o2fillcosttotal;
extern int numdivers;
extern int heo2display;
extern double oxygenfraction, oxygenfractionlast;
extern double hemix, n2mix, hemixlast, n2mixlast;
extern int bailout_breathable;
extern int bailout, bailoutlast;
extern int bailoutpointfracdec[110], bailoutpointfractiondepth[10], bailoutpoint[10][100];
extern double ppo2_limit_lower, ppo2_limit_upper;
extern int optimiseo2, automaticmode, autofinish, gas_table;

extern void setinitial(void);
extern int tissueupdate(int temp_ascent);
extern void tissuetemptransfer(void);
extern void pambtolcalc(int temp_ascent);
extern void pambtolcalctrue(void);
extern double acalc(int i, int n2only, int tempcalc);
extern double bcalc(int i, int n2only, int tempcalc);
extern int getdecompressiontime(void);
extern void storedivedepthandtime();
extern int dispaydepthdata(void);
extern void depthplot(void);
extern void plotdepthdata(void);
extern void decomcalc(void);
void settitletext(void);
void setaxistext(void);
void setmicroaxistext(void);
void titleprofilegraph(int, double, double );
void borderdraw(void);
void printdivehistory(void);
void screenprint(int,int);
extern int licenseread(void);
extern int licensecheck(void);
extern void licensefeetmodewrite(void);
int getgas(int,int,double);
extern double stoptimetisscalc( int, int, double);
extern double tisscalc(int , double, int );
extern void ppo2print(double);
extern void setsafetyfactor(void);
extern void setatmospheric(void);
extern void setcylinderfill(void);
extern void setmissionstart(void);
void settitletextlarge(void);
extern int getdiscsaveddive(void);
extern void savediveondisc2(void);
extern void ppo2exposuretime(void);
extern void tissuegraph(void);
extern void divelist(void);
extern void prodivemenu(void);
extern long fileinfo( struct find_t *find );
extern void missiontotalupdate( double timeinc );
extern void flytolupdate(void);
extern void fractioncalcs(int);
extern void pubub(void);
extern void bubbles( void);
extern void ibhalffile(void);
extern void drawbackground(void);
extern void backgroundtoggle( void);
extern void open_output( void);
extern void activate_graphic_mode( void);
extern void send_linefeed( void);
extern void restore_linefeed( void);
extern void delay1sec(void);
extern void activate_lasergraphic_mode( void);
extern void activate_lasergraphic_mode75( void);
extern void deactivate_lasergraphic_mode( void);
extern void laserleftmargin( void);
extern void lasernumberofbytes( int bytenum);
extern void laserposition( int , int );
extern void algorithmcalc(int i, double time, int fi);
extern void tissuetotissuetemptransfer(void);
extern char *timestr( unsigned t, char *buf );
extern char *datestr( unsigned d, char *buf );
extern void tissueorgtransfer(void);
double ascenttime (double deepest_depth);
double ascenttimediff (double deepest_depth, double depth);
void printtoprinter(int j, int filetext);
void gasfillcalcs(
     int i,int j,int c,int k,int ii,int opxos, unsigned char title[100], double sftemp);
void gasfillsummary(
     int i,int j,int c,int k,int ii,int oppos, unsigned char title[100], double sftemp);
void numgas_calc(int j, int i);
void tradecylinderfill(void);
void tradegasprice(void);
void tradegascurrency(void);
double fillgetgas(void);
extern void helpscreen(int);
int putdisc_gasdata(int row);
void getdisc_gasdata(void);
void printtradecylinderfill(
  int i, int c, int heinc, int topup, int entergaspercent,
  double sftemp, double fillpressure, double workingdepth, double workingppo2, double o2frac, double n2frac, double hefrac, double narcpressure, double hefill, double airfill, double o2fill, double n2fill,
  char message[MAXSTR],
  unsigned char title[100], unsigned char row,
  double hefillpressure, double o2fillpressure, double hefillfrac, double o2fillfrac, double n2fillfrac, double airtopoff,
  int logsave
  );
void tradegasfillcalcs(int row);
int getchyn( void);
int getchyn_defaulty( void);
char *cgetsn( char *buffer, char cn[10], char *default_string );
char *cgetsa( char *buffer );
char *cgetsb( char *buffer );
void savefillcosts(void);
void graphicscreenprint( void);
void seto2optimise(void);
int setoxygenfraction(double ppo2depth);
int dispaydepthdatagas(void);
void putdisc_gasmixdata(char *mix_file);
void getdisc_gasmixdata(char *mix_file);
void tissupdate(int stopsprocessed);
void ttissupdate(int stopsprocessed);

extern FILE *fp, *fprn;

extern float stoplookup_factor[3][4];
  /* = {

  0.0, 3.0, 6.0, 9.0,
  0.0, 4.5, 6.0, 9.0,
  0.0, 6.0, 6.0, 9.0
  };*/

void printdivehistory(void)
{
int i, j, c, k, ii;
int gppos=10, vppos;
unsigned char title[100];
double sftemp=0.00;

  if(he) gppos=6;
  if(ppo2) gppos=6;
  if(ppo2&&he) gppos=3;
  for(j=0; j<divenumber; j++) {

    for(k=14;k<30;k++) {
      _settextposition( k,3);
      sprintf(stitle, "                                                                           ");
      _outtext( stitle) ;
    }

    totaltimepoint[j][0] = timepointa[j][0] +timepointb[j][0];
    number_gasses=0;
    for(k=0; k<NUMGASMIX ;k++) {
      gasmix[j][k][0] = gasmix[j][k][1] = gasmixbartime[j][k] = 0.00;
    }
    gasmix[j][0][0] = nitrogenpoint[j][0];
    gasmix[j][0][1] = heliumpoint[j][0];
    gasmixbartime[j][0] = ( (depthpoint[j][0] + 10.00) / 10.00 ) * (timepointa[j][0] + timepointb[j][0]);
    for(i=1 ;(i<6) && timepointb[j][i]; i++) {
      totaltimepoint[j][i] =
		 totaltimepoint[j][i-1] + timepointa[j][i] +timepointb[j][i];
      numgas_calc(j,i);
    }
    for( ;(i<6) ; i++) {
      totaltimepoint[j][i] = totaltimepoint[j][i-1];
    }
    for(i=6 ;i<(numberpoints[j]); i++) {
      totaltimepoint[j][i] =
		 totaltimepoint[j][i-1] + timepointa[j][i] +timepointb[j][i] +timepointc[j][i];
      // if(i<(numberpoints[j]-1))
      numgas_calc(j,i);
    }
    if(j==0 && !depthpoint[0][0]) {
      if(!depthpoint[1][0]) return;
      j++;
    }

    _clearscreen( _GCLEARSCREEN);
    if( vc.numcolors > 2) _setcolor(7);
    x=4/pixfact; y=8/piyfact;
      _rectangle( _GBORDER, vc.numxpixels-4/pixfact, (480/piyfact)-8/piyfact, x, y );
    _moveto_w( (double)4.00/(double)pixfact, 232.00/(double)piyfact);
      _lineto_w( (double)vc.numxpixels-4.00/(double)pixfact, 232.00/(double)piyfact );
    _moveto_w( (double)348.00/(double)pixfact, 8.00/(double)piyfact);
      _lineto_w( (double)348.00/(double)pixfact, 232.00/(double)piyfact );
    _settextposition(2,2);
    sprintf(stitle,	"Mission %s %-8s Dive %1d %s", divetitl, datebuf, j+1, (ppo2 ? (he ? "DIL=He/N2" : "DIL=N2") : "") );
    _outtext( stitle) ;
    _settextposition(3,3);
    sprintf(stitle,     "Atmospheric pressure= %4dmBar",atmosphericdive);
    _outtext( stitle) ;
      //for(i=0, k=3, vppos=4;(i<3) && timepointb[j][i]; i++, k=k+14) {
      for(i=0, k=3;(i<6) && timepointb[j][i]; i++, k=k+14) {
	vppos= 5;
	if(i>=3) vppos= 10;
	else if(ppo2) vppos=4;
	if(i==3) k=3;
	_settextposition( vppos++, k);
	sprintf(stitle,"Depth%d=%3d%s", i+1, (int)(depthpoint[j][i] * feetfactor+0.49), form);
	_outtext( stitle);
	_settextposition( vppos++, k);
	sprintf(stitle,"Time=%3.0fmins",timepointb[j][i]);
	_outtext( stitle);
	_settextposition( vppos++, k);
	if(ppo2 && !bailoutpoint[j][i]) {
	  if(he) {
	    sprintf(stitle, "%2.0f%%He ", (hemixpoint[j][i]*100.00) );
	    _outtext( stitle) ;
	    sprintf(stitle, "%2.0f%%O2", (O2MIXPOINTJI*100.00) );
	    _outtext( stitle) ;
	    _settextposition( vppos++, k);
	  }
	  else {
	    sprintf(stitle, "(Dil=N2) ");
	    _outtext( stitle) ;
	    _settextposition( vppos++, k);
	  }
	  ppo2print(ppo2point[j][i]);
	  sprintf(stitle, "PPO2=%1d.%02dbar", ppi, ppf);
	  _outtext( stitle) ;

	}
	else {
	  if(bailoutpoint[j][i]) {
	    sprintf(stitle, "!! Bailout !!");
	    _outtext( stitle) ;
	    _settextposition( vppos++, k);
	  }
	  if(air || n2) {
	    if(heo2display) sprintf(stitle, "%2d%%O2 ", (int)(OXYGENPOINTJI*100.00+0.49) );
	    else sprintf(stitle, "%2d%%N2 ", (int)(nitrogenpoint[j][i]*100.00+0.49) );
	    _outtext( stitle) ;
	  }
	  if(he) {
	    sprintf(stitle, "%2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
	    _outtext( stitle) ;
	  }
	}
	_settextposition( vppos++, k);
	sprintf(stitle,"RT=%3.0fmins",totaltimepoint[j][i]);
	_outtext( stitle);
      }
    i=6;
    _settextposition( 16, 3);
    if( vc.numcolors > 2)_settextcolor(3);
    sprintf(title, "******************   DIVE %2d DECOMPRESSION REQUIREMENTS   ******************",j+1);
    _outtext(title);
    for(ii=0, i=6; i<(numberpoints[j]-1); ii+=12) {
      for( ;i<(numberpoints[j]-1) && (11+i-ii)<29 ; i++) {
	if(i>(numberpoints[j]-5) && !stoplookup_factor[sixstopmodedive][numberpoints[j]-1-i]) {
	  continue;
	}
	_settextposition( (11+i-ii), gppos);
	screenprint(i,j);
      }
      if(i<(numberpoints[j]-1) && (11+i-ii)<30 ) {
	_settextposition( (11+i-ii), gppos);
	if( vc.numcolors > 2)_settextcolor(7);
	sprintf(stitle,"press any key for more....");
	_outtext( stitle);
	if( vc.numcolors > 2)_settextcolor(3);
	getch();
	for(k=17;k<30;k++) {
	  _settextposition( k,3);
	  sprintf(stitle, "                                                                           ");
	  _outtext( stitle) ;
	}
      }
    }

    if( vc.numcolors > 2)_settextcolor(7);
    _settextposition(2,46);
    sprintf(stitle,	"Total time to surface=%ldmins ", totaltimetosurface[j]);
    _outtext( stitle) ;
    _settextposition(3,46);
    sprintf(stitle,	"Time to first stop=%ldmins    ", timetofirststop[j]);
    _outtext( stitle) ;
    _settextposition(4,46);
    if( (ppo2cnsmax[j]>100.00) && (vc.numcolors > 2) )_settextcolor(4);
    sprintf(title, "CNS dose: %d%%peak, %d%%dive end", (int)ppo2cnsmax[j], (int)ppo2cns[j]);
    _outtext(title);
    if( vc.numcolors > 2)_settextcolor(7);
    ppo2print( maxppo2[j]);
    _settextposition(5,46);
    if( (maxppo2[j]>2.00) && (vc.numcolors > 2) )_settextcolor(4);
    sprintf(stitle,     "MaxPPO2=%1d.%02dbar", ppi, ppf);
    _outtext( stitle) ;
    if( vc.numcolors > 2)_settextcolor(7);
    sprintf(stitle,     "  OTU=%dunits", (int)diveotu[j]);
    _outtext( stitle) ;
    _settextposition(6,46);
    sprintf(stitle,     "Dive Start:   Day%2d   Time %02d:%02d", divestart[j][0], divestart[j][1], divestart[j][2]);
    _outtext( stitle) ;
    _settextposition(7,46);
    sprintf(stitle,     "Dive Finish:  Day%2d   Time %02d:%02d", divefinish[j][0], divefinish[j][1], divefinish[j][2]);
    _outtext( stitle) ;
    _settextposition(8,46);
    if((j+1)==divenumber) {
      sprintf(stitle,   "Flight Time:  Day%2d   Time %02d:%02d", flytol[0], flytol[1], flytol[2]);
      _outtext( stitle) ;
    }
    i++;
    _settextposition(9,46);
    if(heo2display) sprintf(stitle,	"Surface gas=%2d%%O2", (int)(OXYGENPOINTJI*100.00+0.49) );
    else sprintf(stitle,	 "Surface gas=%2d%%N2", (int)(nitrogenpoint[j][i]*100.00+0.49) );
    _outtext( stitle) ;
    if(he ) {
      sprintf(stitle, ", %2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
      _outtext( stitle) ;
    }
    _settextposition(10,46);
    if((j+1)==divenumber) {
      sprintf(stitle,   "Mission OTU total=%dunits", (int)missionotu);
    }
    else sprintf(stitle, "Surface interval= %4.2fmins",(double)( (int)timepointb[j][i] ) );
    _outtext( stitle) ;
    _settextposition(11,46);
    sprintf(stitle,     "Safety factor= %2d%%",safetyfactordive);
    _outtext( stitle) ;

    if(!ppo2) {
      _settextposition( 29, 20);
      if( vc.numcolors > 2)_settextcolor(7);
      sprintf(stitle,"Do you require gas fill details? Y/<N> ");
      _outtext( stitle);
      c=getchyn();
      if( c=='y' || c=='Y' ) {
	if(gasreservefraction[j][0]) {
	  _settextposition( 29, 20);
	  if( vc.numcolors > 2)_settextcolor(7);
	  sprintf(stitle,"Do you wish to enter NEW gas fill details? Y/<N> ");
	  _outtext( stitle);
	  c=getchyn();
	  if( c=='y' || c=='Y' ) {
	    gasfillcalcs(i,j,c,k,ii,gppos,title,sftemp);
	  }
	  else {
	    gasfillsummary(i,j,c,k,ii,gppos,title,sftemp);
	  }
	}
	else {
	  gasfillcalcs(i,j,c,k,ii,gppos,title,sftemp);
	}
      }
    }

    if(!toln2only) {
      _settextposition(29,5);
      sprintf(stitle,   "He/N2 tol");
      _outtext( stitle) ;
    }
    _settextposition(29,20);
    sprintf(stitle,     "  *****   Send to printer?   Y/<N>  ***** ");
    _outtext( stitle) ;

    c = getchyn();
    if(c=='y' || c=='Y') {
      printtoprinter(j, 0);
    }
    _settextposition(29,20);
    sprintf(stitle,	"  *****   Send to text file %s.txt?   Y/<N>  ***** ",divetitl);
    _outtext( stitle) ;

    c = getchyn();
    if(c=='y' || c=='Y') {
      printtoprinter(j, 1);
    }
  }

}

void screenprint(int i,int j)
{
unsigned char title[100];
double temp;

  if(modf(depthpoint[j][i], &temp)>0.2 && feetfactor==1.00) sprintf(title, "%5.1f%s%cstop: %3.0fmins: ", (depthpoint[j][i] * feetfactor+0.00000), form,( (i==(numberpoints[j]-2) && sixstopmodedive ) ? '*' : ' ' ),timepointb[j][i]+timepointc[j][i]);
  else sprintf(title, "%5d%s%cstop: %3.0fmins: ", (int)(depthpoint[j][i] * feetfactor+0.00000), form,( (i==(numberpoints[j]-2) && sixstopmodedive ) ? '*' : ((depthpoint[j][i]-depthpoint[j][i+1])>4 ? '^' : ' ') ),timepointb[j][i]+timepointc[j][i]);
    _outtext(title);
  if(ppo2 && !bailoutpoint[j][i]) {
    if(he) {
      sprintf(stitle, "(Dil=%3.0f%%He ", (hemixpoint[j][i]*100.00) );
      _outtext( stitle) ;
      sprintf(stitle, "%3.0f%%O2) ", (O2MIXPOINTJI*100.00) );
      _outtext( stitle) ;
    }
    else {
      sprintf(stitle, "(Dil=N2) ");
      _outtext( stitle) ;
    }
    ppo2print(ppo2point[j][i]);
    sprintf(title, "PPO2=%1d.%02dbar", ppi, ppf);
    _outtext(title);
  }
  else {
    if(bailoutpoint[j][i]) {
      sprintf(stitle, "!! Bailout !!        ");
      _outtext( stitle) ;
    }
    if(air || n2) {
      if(heo2display) sprintf(title, "%2d%%O2 ", (int)(OXYGENPOINTJI*100.00+0.49) );
      else sprintf(title, "%2d%%N2 ", (int)(nitrogenpoint[j][i]*100.00+0.49) );
      _outtext(title);
    }
    if(he) {
      sprintf(title, "%2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
      _outtext(title);
    }
  }
  sprintf(title, "  Run Time=%3.0fmins",totaltimepoint[j][i]);
  _outtext(title);
  //sprintf(title, "  Deco Time=%3.0fmins",totaltimepoint[j][i]-totaltimepoint[j][5]);
  //_outtext(title);
}

void printtoprinter(int j, int filetext)
{
int i,k;
unsigned char c, title[100];
double temp;

      if(!filetext) while ((fprn=fopen( "PRN", "w"))==NULL) {
	_settextposition(29,14);
	sprintf(stitle, "  *****   PRINTER DISCONNECTED  Retry?  Y/<N>   *****          ");
	_outtext( stitle) ;
	c = getchyn();
	if(c!='y' && c!='Y') return;
	/*
	strcpy( message, "Unable to open printer file.\n");
	write(fileno(stdout), message, strlen(message));
	exit(3);
	*/
      }
      else {
	strcpy(title, divetitl);
	strcat(title, ".txt");

	if((fprn=fopen( title, "w+"))==NULL) {
	  _settextposition(29,14);
	  sprintf(stitle, "  *****   Cannot open file  Y/<N>   *****          ");
	  _outtext( stitle) ;
	  exit(3);
	}

      }
      fprintf(fprn,    "\n********************************************************************************");
      fprintf(fprn,    "\n                   MISSION %s  %-8s  Dive number:%d %s  ", divetitl, datebuf, j+1, (ppo2 ? (he ? "DIL=He/N2" : "DIL=N2") : "") );
      fprintf(fprn,    "\n********************************************************************************");
      fprintf(fprn,    "\n\n            Atmospheric pressure= %4dmBar    Safety factor=%d%%    \n", atmosphericdive, safetyfactordive);
	for(i=0 ;(i<6) && timepointb[j][i]; i++) {
	  fprintf(fprn,"\n  Depth%d= %3d%s: %3.0fmins ", i+1, (int)(depthpoint[j][i] * feetfactor+0.49), form,timepointb[j][i]);
	  if(ppo2 && !bailoutpoint[j][i]) {
	    if(he) {
	      fprintf(fprn, "(Dil=%2.0f%%He ", (hemixpoint[j][i]*100.00) );
	      fprintf(fprn, "%2.0f%%O2) ", (O2MIXPOINTJI*100.00) );
	    }
	    else fprintf(fprn, "(Dil=N2) ");
	    ppo2print(ppo2point[j][i]);
	    fprintf(fprn, "PPO2=%1d.%02dbar", ppi, ppf);
	  }
	  else {
	    if(bailoutpoint[j][i]) fprintf(fprn, " ! Bailout ! ");
	    if(air || n2) {
	      if(heo2display) fprintf(fprn, "%2d%%O2 ", (int)(OXYGENPOINTJI*100.00+0.49) );
	      else fprintf(fprn, "%2d%%N2 ", (int)(nitrogenpoint[j][i]*100.00+0.49) );
	    }
	    if(he) fprintf(fprn, "%2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
	  }
	  fprintf(fprn, "  Run Time=%3.0fmins",totaltimepoint[j][i]);
	}
      fprintf(fprn,  "\n\n   DIVE %d DECOMPRESSION REQUIREMENTS	      ",j+1);
	for(i=6 ;i<(numberpoints[j]-1); i++) {
	  if(i>(numberpoints[j]-5) && !stoplookup_factor[sixstopmodedive][numberpoints[j]-1-i]) {
	     continue;
	  }
	  if(modf(depthpoint[j][i], &temp)>0.2 && feetfactor==1.00) fprintf(fprn,"\n   %cStop=%4.1f%s: %3.0fmins ",( (i==(numberpoints[j]-2) && sixstopmodedive ) ? '*' : ' ' ), (depthpoint[j][i] * feetfactor+0.0000), form,
		timepointb[j][i]+timepointc[j][i]);
	  else fprintf(fprn,"\n   %cStop=%4d%s: %3.0fmins ",( (i==(numberpoints[j]-2) && sixstopmodedive ) ? '*' : ' ' ), (int)(depthpoint[j][i] * feetfactor+0.0000), form,
		timepointb[j][i]+timepointc[j][i]);
	  if(ppo2 && !bailoutpoint[j][i]) {
	    if(he) {
	      fprintf(fprn, "(Dil=%2.0f%%He ", (hemixpoint[j][i]*100.00) );
	      fprintf(fprn, "%2.0f%%O2) ", (O2MIXPOINTJI*100.00) );
	    }
	    else fprintf(fprn, "(Dil=N2) ");
	    ppo2print(ppo2point[j][i]);
	    fprintf(fprn,"PPO2=%1d.%02dbar", ppi, ppf);
	  }
	  else {
	    if(bailoutpoint[j][i]) fprintf(fprn, " ! Bailout ! ");
	    if(air || n2) {
	      if(heo2display) fprintf(fprn, "%2d%%O2 ", (int)(OXYGENPOINTJI*100.00+0.49) );
	      else fprintf(fprn, "%2d%%N2 ", (int)(nitrogenpoint[j][i]*100.00+0.49) );
	    }
	    if(he && !ppo2) fprintf(fprn,"%2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
	  }
	  fprintf(fprn, "  Run Time=%3.0fmins",totaltimepoint[j][i]);
	}
      fprintf(fprn,  "\n\n      Total time to surface=%ldmins               ", totaltimetosurface[j]);
      fprintf(fprn,    "\n      Time to first stop=%ldmins                  ", timetofirststop[j]);
      fprintf(fprn,    "\n      CNS exposure: %d%%peak, %d%%dive end        ", (int)ppo2cnsmax[j], (int)ppo2cns[j]);
      if( (ppo2cnsmax[j]>100.00) && (vc.numcolors > 2) ) fprintf(fprn,"  !! CNS WARNING  !!");
      fprintf(fprn,    "\n      CNS exposure=%3d%%                           ",(int)ppo2cns[j]);
      ppo2print( maxppo2[j]);
      fprintf(fprn,    "\n      Max PPO2=%1d.%02dbar  OTU=%dunits", ppi, ppf, (int)diveotu[j]);
      if((j+1)==divenumber) fprintf(fprn,"  OTUtotal=%d", (int)missionotu);
      if( (maxppo2[j]>2.00) && (vc.numcolors > 2) ) fprintf(fprn,"   !! PPO2 WARNING !!");
      i++;
      fprintf(fprn,    "\n      Dive Start:   Day%2d   Time %02d:%02d         ", divestart[j][0], divestart[j][1], divestart[j][2]);
      fprintf(fprn,    "\n      Dive Finish:  Day%2d   Time %02d:%02d         ", divefinish[j][0], divefinish[j][1], divefinish[j][2]);
      if((j+1)==divenumber) fprintf(fprn,"\n      Flight Time:  Day%2d   Time %02d:%02d", flytol[0], flytol[1], flytol[2]);
      if(heo2display) fprintf(fprn,	 "\n      Surface gas=%2d%%O2", (int)(OXYGENPOINTJI*100.00+0.49) );
      else fprintf(fprn,	 "\n      Surface gas=%2d%%N2", (int)(nitrogenpoint[j][i]*100.00+0.49) );
      if(he ) fprintf(fprn, ", %2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
      fprintf(fprn,    "\n      Surface interval= %3.0fmins                     ",(double)( (int)timepointb[j][i] ) );
      if(gasreservefraction[j][0]) {
	fprintf(fprn, "\n\n   DIVE %d GAS FILLING DETAILS",j+1);
	fprintf(fprn, "\n\n   Number of mixes=%2d  Breathing rate=%3.2f%s ",number_gasses+1, breathingratedive/cuft_ltr_factor, cuftorltrmin);
	for(k=0; k<(number_gasses+1); k++) {
	  fprintf(fprn, "\n\n    Gas mix%d: ", k+1);
	  fprintf(fprn, "Reserve= %.0f%%",gasreservefraction[j][k]*100.0);
	  if(feetfactor==1.00) {
	    fprintf(fprn, "Cylinder= %.0flitres",cylindersize[j][k]);
	  }
	  else {
	    fprintf(fprn, "Cylinder water capacity= %.2fcuft",cylindersize[j][k]/cuft_ltr_factor);
	  }
	  if(filltotal[j][k] > maxcylinderpressure[j][k]) {
	    fprintf(fprn, "\n    WARNING: Fill pressure greater than maximum allowed for cylinder!");
	    fprintf(fprn, "\n             DO NOT FILL TO THE PRESSURE GIVEN BELOW");
	    fprintf(fprn, "\n             Recalculate gas usage or DO NOT PERFORM DIVE");
	  }
	  if(air || n2 && !ppo2) {
	    if(heo2display) fprintf(fprn, "\n    %2d%%O2 ", (int)(GASMIXO2*100.00+0.49) );
	    else fprintf(fprn, "\n    %2d%%N2 ", (int)(gasmix[j][k][0]*100.00+0.49) );
	  }
	  if(he && !ppo2) {
	    fprintf(fprn, "%2d%%He", (int)(gasmix[j][k][1]*100.00+0.49) );
	  }
	  fprintf(fprn, "  Fill pressure=%ld%s ",(long)( filltotal[j][k] * psifactor+0.49 ),porb);
	  fprintf(fprn, "(Free gas volume at 1bar=%.1f%s)", filltotal[j][k]*cylindersize[j][k]/cuft_ltr_factor, cuftorltr);
	}
      }
      if(!toln2only) {
	fprintf(fprn,   "\n\nHe/N2 tol");
      }
      fprintf(fprn,    "\n********************************************************************************");
      if(filetext) {
	for(i=0 ;(i<6) && timepointb[j][i]; i++) {
	  fprintf(fprn,"\n  Depth%d= %3d%s: %3.0fmins ", i+1, (int)(depthpoint[j][i] * feetfactor+0.49), form,timepointb[j][i]);
	  if(ppo2 && !bailoutpoint[j][i]) {
	    if(he) {
	      fprintf(fprn, "(Dil=%2.0f%%He ", (hemixpoint[j][i]*100.00) );
	      fprintf(fprn, "%2.0f%%O2) ", (O2MIXPOINTJI*100.00) );
	    }
	    else fprintf(fprn, "(Dil=N2) ");
	    ppo2print(ppo2point[j][i]);
	    fprintf(fprn, "PPO2=%1d.%02dbar", ppi, ppf);
	  }
	  else {
	    if(bailoutpoint[j][i]) fprintf(fprn, " ! Bailout ! ");
	    if(air || n2) {
	      if(heo2display) fprintf(fprn, "%2d%%O2 ", (int)(OXYGENPOINTJI*100.00+0.49) );
	      else fprintf(fprn, "%2d%%N2 ", (int)(nitrogenpoint[j][i]*100.00+0.49) );
	    }
	    if(he && !ppo2) fprintf(fprn, "%2d%%He", (int)(heliumpoint[j][i]*100.00+0.49) );
	  }
	  fprintf(fprn, "  Run Time=%3.0fmins",totaltimepoint[j][i]);
	}
	fprintf(fprn,"\n  Depth,    Gas,  Stop,  RT");
	for(i=(numberpoints[j]-2); i>5 ;i--) {
	  if(i>(numberpoints[j]-5) && !stoplookup_factor[sixstopmodedive][numberpoints[j]-1-i]) {
	     continue;
	  }
	  if(modf(depthpoint[j][i], &temp)>0.2 && feetfactor==1.00) fprintf(fprn,"\n   %4.1f, ", (depthpoint[j][i] * feetfactor+0.0000));
	  else fprintf(fprn,"\n   %4d,  ", (int)(depthpoint[j][i] * feetfactor+0.0000));
	  if(ppo2 && !bailoutpoint[j][i]) {
	    if(he) {
		fprintf(fprn, "(Dil=%2.0f%%He ", (hemixpoint[j][i]*100.00) );
		fprintf(fprn, "%2.0f%%O2) ", (O2MIXPOINTJI*100.00) );
	    }
	    else fprintf(fprn, "(Dil=N2) ");
	    ppo2print(ppo2point[j][i]);
	    fprintf(fprn,"PPO2=%1d.%02dbar", ppi, ppf);
	  }
	  else {
	    if(air || n2) {
	      if(heo2display) fprintf(fprn, "%2d", (int)(OXYGENPOINTJI*100.00+0.49) );
	      else fprintf(fprn, "/%2d", (int)(nitrogenpoint[j][i]*100.00+0.49) );
	    }
	    if(he) fprintf(fprn,"/%2d", (int)(heliumpoint[j][i]*100.00+0.49) );
	  }
	  fprintf(fprn,",   %3.0f",timepointb[j][i]+timepointc[j][i]);
	  fprintf(fprn, ",   %3.0f",totaltimepoint[j][i]);
	}
      }
      else fprintf(fprn,	"%c",12);
      fclose( fprn);

}

void gasfillcalcs(
int i,int j,int c,int k,int ii,int gppos, unsigned char title[100], double sftemp)
{
int kn;
	for(k=16;k<30;k++) {
	  _settextposition( k,3);
	  sprintf(stitle, "                                                                           ");
	  _outtext( stitle) ;
	}
	if( vc.numcolors > 2)_settextcolor(3);
	_settextposition( 16, 3);
	   sprintf(title, "DIVE %2d GAS FILLING DETAILS  Number of mixes=%2d  Breathing rate=%3.2f%s ",j+1,number_gasses+1, breathingratedive/cuft_ltr_factor, cuftorltrmin);
	_outtext(title);
	if( vc.numcolors > 2)_settextcolor(7);
	for(k=0, kn=0; k<(number_gasses+1); k++, kn++) {
	  if(k==6) {
	    for(k=17;k<30;k++) {
	      _settextposition( k,3);
	      sprintf(stitle, "                                                                           ");
	      _outtext( stitle) ;
	    }
	    k=6;
	    kn=0;
	  }
	  _settextposition( 17+2*kn,4);
	  sprintf(stitle, "Gas mix%d:", k+1);
	  _outtext( stitle) ;
	  _settextposition( 18+2*kn, 4);
	  if(air || n2 && !ppo2) {
	    if(heo2display) sprintf(stitle, "%2d%%O2 ", (int)(GASMIXO2*100.00+0.49) );
	    else sprintf(stitle, "%2d%%N2 ", (int)(gasmix[j][k][0]*100.00+0.49) );
	    _outtext( stitle) ;
	  }
	  if(he && !ppo2) {
	    sprintf(stitle, "%2d%%He", (int)(gasmix[j][k][1]*100.00+0.49) );
	    _outtext( stitle) ;
	  }
	  do {
	    _settextposition( 17+2*kn,17);
	    sprintf(stitle, "Reserve= __%%");
	    _outtext( stitle) ;
	    _settextposition( 17+2*kn,26);
	    cnumbuf[0]=3;
	    numbuf = cgetsn( cnumbuf, "", "" );
	    if(!*numbuf) return;
	    sftemp = ( (double)atoi( numbuf ) ) ;
	  } while( (sftemp < 0.00) || (sftemp > 90.00) || !*numbuf);
	  gasreservefraction[j][k] = sftemp/100.00;
	  if(!gasreservefraction[j][k]) gasreservefraction[j][k] = 0.0001;
	  _settextposition( 17+2*kn,17);
	  sprintf(stitle, "Reserve= %.0f%%           ",gasreservefraction[j][k]*100.0);
	  _outtext( stitle) ;
	  do {
	    _settextposition( 17+2*kn,31);
	    sprintf(stitle, "Maximum working cylinder pressure ____%s",porb);
	    _outtext( stitle) ;
	    _settextposition( 17+2*kn,65);
	    cnumbuf[0]=5;
	    numbuf = cgetsn( cnumbuf, "", "" );
	    if(!*numbuf) return;
	    sftemp = ( (double)atof( numbuf ) ) ;
	  } while( (sftemp < 0.00) || !*numbuf);
	  maxcylinderpressure[j][k] = sftemp/psifactor;
	  if(feetfactor==1.00) {
	    do {
	      _settextposition( 17+2*kn,31);
	      sprintf(stitle, "Cylinder size = ____litres                  ");
	      _outtext( stitle) ;
	      _settextposition( 17+2*kn,47);
	      cnumbuf[0]=4;
	      numbuf = cgetsn( cnumbuf, "", "" );
	      if(!*numbuf) return;
	      sftemp = ( (double)atof( numbuf ) ) ;
	    } while( (sftemp < 0.00) || !*numbuf);
	    cylindersize[j][k] = sftemp;
	    _settextposition( 17+2*kn,31);
	    sprintf(stitle, "Cylinder= %.0flitres                       ",cylindersize[j][k]);
	    _outtext( stitle) ;
	  }
	  else {
	    do {
	      _settextposition( 17+2*kn,31);
	      sprintf(stitle, "Free air capacity of cylinder = ___cubic feet");
	      _outtext( stitle) ;
	      _settextposition( 17+2*kn,63);
	      cnumbuf[0]=4;
	      numbuf = cgetsn( cnumbuf, "", "" );
	      if(!*numbuf) return;
	      sftemp = ( (double)atof( numbuf ) ) ;
	    } while( (sftemp < 0.00) || !*numbuf);
	    freecylindersize[j][k] = sftemp*cuft_ltr_factor;
	    cylindersize[j][k] = freecylindersize[j][k] / maxcylinderpressure[j][k];
	    _settextposition( 17+2*kn,31);
	    sprintf(stitle, "Cylinder water capacity= %.2fcuft              ",cylindersize[j][k]/cuft_ltr_factor);
	    _outtext( stitle) ;
	  }
	  _settextposition( 18+2*kn, 17);
	  /*
	  sprintf(stitle, "Gas bar mins=%.1f, ",gasmixbartime[j][k]);
	  _outtext( stitle) ;
	  */
	  filldive[j][k] = (gasmixbartime[j][k]*breathingratedive)/cylindersize[j][k];
	  fillres[j][k] = filldive[j][k]*gasreservefraction[j][k]/(1.00-gasreservefraction[j][k]);
	  filltotal[j][k] = filldive[j][k] + fillres[j][k];
	  if(filltotal[j][k] > maxcylinderpressure[j][k])
	    if( vc.numcolors > 2)_settextcolor(4);
	  else
	    if( vc.numcolors > 2)_settextcolor(15);
	  sprintf(stitle, "Fill pressure= %ld%s ",(long)( filltotal[j][k] * psifactor+0.49 ),porb);
	  _outtext( stitle) ;
	  if( vc.numcolors > 2)_settextcolor(7);
	  sprintf(stitle, "(Free gas volume at 1bar=%.1f%s)", filltotal[j][k]*cylindersize[j][k]/cuft_ltr_factor, cuftorltr);
	  _outtext( stitle) ;
	}
}

void gasfillsummary(
int i,int j,int c,int k,int ii,int gppos, unsigned char title[100], double sftemp)
{
int kn;
	for(k=16;k<30;k++) {
	  _settextposition( k,3);
	  sprintf(stitle, "                                                                           ");
	  _outtext( stitle) ;
	}
	if( vc.numcolors > 2)_settextcolor(3);
	_settextposition( 16, 3);
	   sprintf(title, "DIVE %2d GAS FILLING DETAILS  Number of mixes=%2d  Breathing rate=%3.2f%s ",j+1,number_gasses+1, breathingratedive/cuft_ltr_factor, cuftorltrmin);
	_outtext(title);
	if( vc.numcolors > 2)_settextcolor(7);
	for(k=0, kn=0; k<(number_gasses+1); k++, kn++) {
	  if(k==6) {
	    _settextposition(29,20);
	    sprintf(stitle,     "  ****  Press any key for more....  **** ");
	    _outtext( stitle) ;
	    c = getch();
	    for(k=17;k<30;k++) {
	      _settextposition( k,3);
	      sprintf(stitle, "                                                                           ");
	      _outtext( stitle) ;
	    }
	    k=6;
	    kn=0;
	  }
	  _settextposition( 17+2*kn,4);
	  sprintf(stitle, "Gas mix%d:", k+1);
	  _outtext( stitle) ;
	  _settextposition( 18+2*kn, 4);
	  if(air || n2 && !ppo2) {
	    if(heo2display) sprintf(stitle, "%2d%%O2 ", (int)(GASMIXO2*100.00+0.49) );
	    else sprintf(stitle, "%2d%%N2 ", (int)(gasmix[j][k][0]*100.00+0.49) );
	    _outtext( stitle) ;
	  }
	  if(he && !ppo2) {
	    sprintf(stitle, "%2d%%He", (int)(gasmix[j][k][1]*100.00+0.49) );
	    _outtext( stitle) ;
	  }
	  _settextposition( 17+2*kn,17);
	  sprintf(stitle, "Reserve= %.0f%%           ",gasreservefraction[j][k]*100.0);
	  _outtext( stitle) ;
	  if(feetfactor==1.00) {
	    _settextposition( 17+2*kn,31);
	    sprintf(stitle, "Cylinder= %.0flitres               ",cylindersize[j][k]);
	    _outtext( stitle) ;
	  }
	  else {
	    _settextposition( 17+2*kn,31);
	    sprintf(stitle, "Cylinder water capacity= %.2fcuft              ",cylindersize[j][k]/cuft_ltr_factor);
	    _outtext( stitle) ;
	  }
	  _settextposition( 18+2*kn, 17);
	  if(filltotal[j][k] > maxcylinderpressure[j][k])
	    if( vc.numcolors > 2)_settextcolor(4);
	  else
	    if( vc.numcolors > 2)_settextcolor(15);
	  sprintf(stitle, "Fill pressure= %ld%s ",(long)( filltotal[j][k] * psifactor+0.49 ),porb);
	  _outtext( stitle) ;
	  if( vc.numcolors > 2)_settextcolor(7);
	  sprintf(stitle, "(Free gas volume at 1bar=%.1f%s)", filltotal[j][k]*cylindersize[j][k]/cuft_ltr_factor, cuftorltr);
	  _outtext( stitle) ;
	}
}

void numgas_calc(int j, int i)
{
int gc;

      for(gc=0; gc<NUMGASMIX && !ppo2 ;gc++) {
	if( (nitrogenpoint[j][i]==gasmix[j][gc][0]) &&  (heliumpoint[j][i] == gasmix[j][gc][1]) ) {
	  if(i!=numberpoints[j]-1) gasmixbartime[j][gc] += ( (depthpoint[j][i] + 10.00) / 10.00 ) * (timepointb[j][i]+timepointc[j][i]);
	  if(depthpoint[j][i]<depthpoint[j][i-1])
	       gasmixbartime[j][gc] += ( ( (depthpoint[j][i-1] - depthpoint[j][i])/2.00 + depthpoint[j][i] + 10.00) / 10.00 ) * (timepointa[j][i]);
	  else
	       gasmixbartime[j][gc] += ( (depthpoint[j][i] + 10.00) / 10.00 ) * (timepointa[j][i]);
	  break;
	}
	else {
	  if( gc==(NUMGASMIX-1) ) {
	    if(timepointa[j][i] || timepointb[j][i]) number_gasses++;
	    else break; /* ignore if no time associated with depth */
	    if(number_gasses==NUMGASMIX ) {
	      sprintf(stitle, "\nToo many gas switches. Gas fill calculations aborted");
	      _outtext(stitle);
	      number_gasses--;
	      break;
	    }
	    else {
	      gasmix[j][number_gasses][0] = nitrogenpoint[j][i];
	      gasmix[j][number_gasses][1] = heliumpoint[j][i];
	      gasmixbartime[j][number_gasses] = ( (depthpoint[j][i] + 10.00) / 10.00 ) * (timepointa[j][i] + timepointb[j][i] + timepointc[j][i]);
	    }
	    break;
	  }
	}
      }
}

void printtradecylinderfill(
  int i, int c, int heinc, int topup, int entergaspercent,
  double sftemp, double fillpressure, double workingdepth, double workingppo2, double o2frac, double n2frac, double hefrac, double narcpressure, double hefill, double airfill, double o2fill, double n2fill,
  char message[MAXSTR],
  unsigned char title[100], unsigned char row,
  double hefillpressure, double o2fillpressure, double hefillfrac, double o2fillfrac, double n2fillfrac, double airtopoff,
  int logsave
  )
{
  int n;

  if(logsave) {
    if((fprn=fopen( "gasfill.txt", "a+"))==NULL) {
      strcpy( message, "Unable to open printer file.\n");
      write(fileno(stdout), message, strlen(message));
      exit(3);
    }
  }
  else {
    if((fprn=fopen( "PRN", "w"))==NULL) {
      strcpy( message, "Unable to open printer file.\n");
      write(fileno(stdout), message, strlen(message));
      exit(3);
    }
  }
  fprintf(fprn,    "\n\n********************************************************************************");
    fprintf(fprn, "\n          Maximum working cylinder pressure %.0f%s", trademaxcylinderpressure*psifactor, porb);
  if(feetfactor==1.00) {
      fprintf(fprn, "\n          Cylinder size = %.1flitres",tradecylindersize);
  }
  else {
      fprintf(fprn, "\n          Free air capacity of cylinder = %.2fcubic feet",tradefreecylindersize/cuft_ltr_factor);
  }

  if(topup)
      fprintf(fprn, "\n          Air top off Required");
  if(heinc)
      fprintf(fprn, "\n          Helium used in mix");
  if(topup && entergaspercent)
      fprintf(fprn, "\n          Current gas entered as %s",entergaspercent ? "Percentages" : "Pressures");
  fprintf(fprn,    "\n");
  if(topup && !entergaspercent && heinc)
      fprintf(fprn, "\n          Current He Fill pressure= %.1f%s",hefillpressure*psifactor,porb);
  if(topup && !entergaspercent)
      fprintf(fprn, "\n          Current O2 Fill pressure= %.1f%s",o2fillpressure*psifactor,porb);
  if(topup && entergaspercent && heinc)
      fprintf(fprn, "\n          Current He Fill percent= %.1f%%",hefillfrac*100.00,porb);
  if(topup && entergaspercent)
      fprintf(fprn, "\n          Current O2 Fill percent= %.1f%%",o2fillfrac*100.00,porb);
  if(topup)
      fprintf(fprn, "\n          Current cylinder pressure= %.1f%s",currentfillpressure*psifactor,porb);
  fprintf(fprn, "\n          Required fill pressure= %.1f%s",fillpressure*psifactor,porb);
  fprintf(fprn, "\n          Final gas mix: ");
  if(o2frac>0.00)
    fprintf(fprn, "%.1f%%O2, ", (o2frac*100.00) );
  if(n2frac>0.00)
    fprintf(fprn, "%.1f%%N2, ", (n2frac*100.00) );
  if(hefrac>0.00)
    fprintf(fprn, "%.1f%%He", (hefrac*100.00) );
  if(!topup) {
    if(heinc) {
      fprintf(fprn, "\n          Helium fill=%.1f%s",hefill*psifactor,porb);
      fprintf(fprn, "\n          Air fill=%.1f%s",airfill*psifactor,porb);
      fprintf(fprn, "\n          Oxygen fill=%.1f%s",o2fill*psifactor,porb);
    }
    else {
      if(o2frac<0.21) {
	fprintf(fprn, "\n          Nitrogen fill=%.1f%s",n2fill*psifactor,porb);
	fprintf(fprn, "\n          Air fill=%.1f%s",airfill*psifactor,porb);
      }
      else {
	fprintf(fprn, "\n          Air fill=%.1f%s",airfill*psifactor,porb);
	fprintf(fprn, "\n          Oxygen fill=%.1f%s",o2fill*psifactor,porb);
      }
    }
  }
  if(!topup) {
    if(heinc) {
      fprintf(fprn, "\n          Oxygen fill: %.1f%s, %.1f%s, %s%.3f", o2fill*psifactor, porb, o2fillltr/cuft_ltr_factor, cuftorltr, gascurrency, o2fillprice);
      fprintf(fprn, "\n          Helium fill: %.1f%s, %.1f%s, %s%.3f", hefill*psifactor, porb, hefillltr/cuft_ltr_factor, cuftorltr, gascurrency, hefillprice);
      fprintf(fprn, "\n          Air fill   : %.1f%s, %.1f%s, %s%.3f", airfill*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
    }
    else {
      if(o2frac<0.21) {
	fprintf(fprn, "\n          Nitrogen fill=%.1f%s",n2fill*psifactor,porb);
	fprintf(fprn, "\n          Air fill   : %.1f%s, %.1f%s, %s%.3f", airfill*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
      }
      else {
	fprintf(fprn, "\n          Oxygen fill: %.1f%s, %.1f%s, %s%.3f", o2fill*psifactor, porb, o2fillltr/cuft_ltr_factor, cuftorltr, gascurrency, o2fillprice);
	fprintf(fprn, "\n          Air fill   : %.1f%s, %.1f%s, %s%.3f", airfill*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
      }
    }
  }
  if(topup) {
    fprintf(fprn, "\n          Air topoff: %.1f%s, %.1f%s, %s%.3f", airtopoff*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
  }

  if(feetfactor=1.00) fprintf(fprn, "\n\n          Date:%c%c/%c%c/%c%c", tradedatebuf[3], tradedatebuf[4], tradedatebuf[0], tradedatebuf[1], tradedatebuf[6], tradedatebuf[7]);
  else                fprintf(fprn, "\n\n          Date:%c%c/%c%c/%c%c", tradedatebuf[0], tradedatebuf[1], tradedatebuf[3], tradedatebuf[4], tradedatebuf[6], tradedatebuf[7]);
  fprintf(fprn,    "\n\n********************************************************************************");
  if(logsave) {
    for(n=0;n<numdivers;n++) {
      fprintf(fprn, "\n\n\nActual as filled gas analysis __________________________");
      fprintf(fprn, "\n\nSigned by %s as correct data for cylinder fill __________________",divername[n]);
    }
    fprintf(fprn,    "\n\n********************************************************************************");
/*
    fprintf(fprn,	 "\n\n             Air            Oxygen         Helium", gascurrency, airfillpricetotal, gascurrency, o2fillpricetotal, gascurrency, hefillpricetotal);
    fprintf(fprn,    "\n\n      Costs: %4s%8.2f   %4s%8.2f   %4s%8.2f", gascurrency, (aircost/airprice)*airfillpricetotal, gascurrency, (o2cost/o2price)*o2fillpricetotal, gascurrency, (hecost/heprice)*hefillpricetotal);
    fprintf(fprn,      "\n      Sales: %4s%8.2f   %4s%8.2f   %4s%8.2f", gascurrency, airfillpricetotal, gascurrency, o2fillpricetotal, gascurrency, hefillpricetotal);
    fprintf(fprn,    "\nOverall Total gas sales = %s%.2f, Total gas costs = %s%.2f",
			gascurrency, (airfillpricetotal+o2fillpricetotal+hefillpricetotal), gascurrency, ((aircost/airprice)*airfillpricetotal+(o2cost/o2price)*o2fillpricetotal+(hecost/heprice)*hefillpricetotal)	);
*/
    fprintf(fprn,    "\n\n                 Costs          Sales");
    fprintf(fprn,    "\n\n    Air:     %4s%8.2f   %4s%8.2f", gascurrency, (aircost/airprice)*airfillpricetotal, gascurrency, airfillpricetotal);
    fprintf(fprn,      "\n    Oxygen:  %4s%8.2f   %4s%8.2f", gascurrency, (o2cost/o2price)*o2fillpricetotal, gascurrency, o2fillpricetotal);
    fprintf(fprn,      "\n    Helium:  %4s%8.2f   %4s%8.2f", gascurrency, (hecost/heprice)*hefillpricetotal, gascurrency, hefillpricetotal);
    fprintf(fprn,    "\n\n    Total:   %4s%8.2f   %4s%8.2f",
			gascurrency, ((aircost/airprice)*airfillpricetotal+(o2cost/o2price)*o2fillpricetotal+(hecost/heprice)*hefillpricetotal),gascurrency, (airfillpricetotal+o2fillpricetotal+hefillpricetotal) );
  }
  else {
    for(n=0;n<numdivers;n++) {
      fprintf(fprn, "\n\nActual as filled gas analysis __________________________");
      fprintf(fprn, "\nSigned by %s as correct data for cylinder fill __________________",divername[n]);
    }
  }

  if( !logsave && lasermode ) {
      fprintf(fprn,"%c",12);
  }
  fclose(fprn);
}


double fillgetgas(void)
{
double tempgasfrac;
      cnumbuf[0]=3;
      numbuf = cgetsn( cnumbuf, "", "" );
      tempgasfrac = ((double)atof( numbuf ))/100.00;
      return tempgasfrac;
}

double pricegetgas(void)
{
double tempgasprice;
      cnumbuf[0]=7;
      numbuf = cgetsn( cnumbuf, "", "" );
      tempgasprice = ((double)atof( numbuf ));
      /*
      if(!*numbuf ) heliumfraction = heliumfractionlast;
      if( !heliumfraction ) heliumfraction = 0.001 ;
      if( (heliumfraction+nitrogenfraction) >= 1.00 ) heliumfraction = 0.001 ;
      */
      return tempgasprice;
}

void tradegasprice(void)
{
int i, c;
double sftemp;
unsigned char row;

  drawbackground();
  _moveto( 75/pixfact, 11/piyfact);
  _outgtext("TRADE GAS PRICE");
  _moveto( 75/pixfact, 26/piyfact);
  _outgtext("       ");
  if( vc.numcolors > 2) _setcolor(7);
  row=8;

	helpscreen(41);
	_settextposition( row++, 20);
	sprintf(stitle, "Helium sell price =%.3f%s\\%s", heprice*cuft_ltr_factor, gascurrency, cuftorltr);
	_outtext( stitle) ;
	_settextposition( row, 20);
	sprintf(stitle, "Enter new Helium sell price =______%s\\%s", gascurrency, cuftorltr);
	_outtext( stitle) ;
	_settextposition( row++, 49);
	sftemp = pricegetgas()/cuft_ltr_factor;
	if(sftemp) heprice = sftemp;
	_settextposition( row++, 20);
	sprintf(stitle, "Helium cost =%.3f%s\\%s", hecost*cuft_ltr_factor, gascurrency, cuftorltr);
	_outtext( stitle) ;
	_settextposition( row, 20);
	sprintf(stitle, "Enter new Helium cost =______%s\\%s", gascurrency, cuftorltr);
	_outtext( stitle) ;
	_settextposition( row++, 43);
	sftemp = pricegetgas()/cuft_ltr_factor;
	if(sftemp) hecost = sftemp;

	helpscreen(41);
	_settextposition( row++, 20);
	sprintf(stitle, "Oxygen sell price =%.3f%s\\%s", o2price*cuft_ltr_factor, gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row, 20);
	sprintf(stitle, "Enter new Oxygen sell price =______%s\\%s", gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row++, 49);
	sftemp = pricegetgas()/cuft_ltr_factor;
	if(sftemp) o2price = sftemp;
	_settextposition( row++, 20);
	sprintf(stitle, "Oxygen cost =%.3f%s\\%s", o2cost*cuft_ltr_factor, gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row, 20);
	sprintf(stitle, "Enter new Oxygen cost =______%s\\%s", gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row++, 43);
	sftemp = pricegetgas()/cuft_ltr_factor;
	if(sftemp) o2cost = sftemp;

	helpscreen(41);
	_settextposition( row++, 20);
	sprintf(stitle, "Air sell price =%.3f%s\\%s", airprice*cuft_ltr_factor, gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row, 20);
	sprintf(stitle, "Enter new air sell price    =______%s\\%s", gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row++, 49);
	sftemp = pricegetgas()/cuft_ltr_factor;
	if(sftemp) airprice = sftemp;
	_settextposition( row++, 20);
	sprintf(stitle, "Air cost =%.3f%s\\%s", aircost*cuft_ltr_factor, gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row, 20);
	sprintf(stitle, "Enter new air cost    =______%s\\%s", gascurrency, cuftorltrnos);
	_outtext( stitle) ;
	_settextposition( row++, 43);
	sftemp = pricegetgas()/cuft_ltr_factor;
	if(sftemp) aircost = sftemp;
	savefillcosts();

  _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );
  x=250/pixfact; y=29/piyfact;
  if( vc.numcolors > 2) {
    _setcolor(7);
    _rectangle( _GFILLINTERIOR, vc.numxpixels-10/pixfact, 44/piyfact, x, y );
    _setcolor(6);
  }
  else {
    _setviewport( x, y, 640/pixfact,44/piyfact );
    _clearscreen( _GVIEWPORT);
    _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );
  }
	//helpscreen(0);
	row = putdisc_gasdata(row);
	row++;
	_settextposition( row++, 20);
	sprintf(stitle, "Press any key to continue......");
	_outtext( stitle) ;
	getch();
}

void tradegascurrency(void)
{
unsigned char row;
  drawbackground();
  _moveto( 75/pixfact, 11/piyfact);
  _outgtext("TRADE GAS CURRENCY");
  _moveto( 75/pixfact, 26/piyfact);
  _outgtext("       ");
  if( vc.numcolors > 2) _setcolor(7);
  helpscreen(42);
  row=8;
    _settextposition( row++, 20);
    sprintf(stitle, "Gas currency= %s", gascurrency);
    _outtext( stitle) ;
    _settextposition( row,20);
    sprintf(stitle, "Enter new gas currency (eg: $) _____");
    _outtext( stitle) ;
    _settextposition( row++,51);
    cnumbuf[0]=6;
    numbuf = cgetsa( cnumbuf );
  if(!*numbuf) return;
  strcpy(gascurrency, numbuf);

  //helpscreen(0);
  row = putdisc_gasdata(row);
  row++;
  _settextposition( row++, 20);
  sprintf(stitle, "Press any key to continue......");
  _outtext( stitle) ;
  getch();

}

void getdisc_gasmixdata(char *mix_file)
{
int i;

int gas[3][NUMGASMIX] = {
   79,	20, 01,  68,  64,  60,	50,  47,  29,  36,
    0,	 0,  0,   0,   0,   0,	 0,  35,  60,  50,
  610,	90, 60, 360, 320, 280, 200, 730, 1300, 980
};

  if( !(fp=fopen( mix_file, "rb" )) ) {
    for(i=0; i<NUMGASMIX; i++) {
      gasmixtable[i][0] = (float)gas[0][i]/100.00;
      gasmixtable[i][1] = (float)gas[1][i]/100.00;
      gasmixtable[i][2] = (float)gas[2][i]/10.00;
      gasstatus[i]=1;
      gasused[i]=0;
    }
    return;
  }

  fread( gasmixtable, sizeof(double), (size_t) 3*NUMGASMIX, fp);
  fread( gasstatus, sizeof(char), (size_t) NUMGASMIX, fp);
  fread( gasused, sizeof(char), (size_t) NUMGASMIX, fp);

  fclose(fp);

}

void putdisc_gasmixdata(char *mix_file)
{
char title[20];
int i;

  strcpy(title, mix_file);
  strcat(title, ".mix");
  if( !(fp=fopen( title, "w+b" )) ) {
    _settextposition(10,10);
    sprintf(stitle, "Cannot save gas mix data to disc. Disc may be full.");
    _outtext( stitle) ;
    return;
  }

  fwrite( gasmixtable, sizeof(double), (size_t) 3*NUMGASMIX, fp);
  fwrite( gasstatus, sizeof(char), (size_t) NUMGASMIX, fp);
  fwrite( gasused, sizeof(char), (size_t) NUMGASMIX, fp);

  fclose(fp);

  strcpy(title, mix_file);
  strcat(title, ".mit");
  if( !(fp=fopen( title, "w" )) ) {
    _settextposition(10,10);
    sprintf(stitle, "Cannot save gas mix data to disc. Disc may be full.");
    _outtext( stitle) ;
    return;
  }
  fprintf(fp,"A");

  for(i=0;i<NUMGASMIX;i++) {
    if(i==0) fprintf(fp, "%03.0f %03.0f %04.0f %1d %c",(float)79.00,(float)0.00,gasmixtable[i][2]*10.00,gasused[i],'B'+i);
    else fprintf(fp, "%03.0f %03.0f %04.0f %1d %c",gasmixtable[i][0]*100.00,gasmixtable[i][1]*100.00,gasmixtable[i][2]*10.00,gasused[i]&gasstatus[i],'B'+i);

  }
  fprintf(fp, "%c",27);
  fclose(fp);

}

void getdisc_gasdata(void)
{

  if( !(fp=fopen( "gasdata.dat", "rb" )) ) {
    heprice=1.00; o2price=1.00; airprice=1.00;
    if(feetfactor==1.00) strcpy( gascurrency, "#" );
    else strcpy( gascurrency, "$" );
    return;
  }

  fread( &heprice, sizeof(double), (size_t) 1, fp);
  fread( &o2price, sizeof(double), (size_t) 1, fp);
  fread( &airprice, sizeof(double), (size_t) 1, fp);
  fread( gascurrency, sizeof(char), (size_t) 9, fp);
  fread( &hecost, sizeof(double), (size_t) 1, fp);
  fread( &o2cost, sizeof(double), (size_t) 1, fp);
  fread( &aircost, sizeof(double), (size_t) 1, fp);

  fclose(fp);

}

int putdisc_gasdata(int row)
{

  if( !(fp=fopen( "gasdata.dat", "w+b" )) ) {
    _settextposition(row++,10);
    sprintf(stitle, "Cannot save gas price data to disc. Disc may be full.");
    _outtext( stitle) ;
     return row;
  }

  fwrite( &heprice, sizeof(double), (size_t) 1, fp);
  fwrite( &o2price, sizeof(double), (size_t) 1, fp);
  fwrite( &airprice, sizeof(double), (size_t) 1, fp);
  fwrite( gascurrency, sizeof(char), (size_t) 9, fp);
  fwrite( &hecost, sizeof(double), (size_t) 1, fp);
  fwrite( &o2cost, sizeof(double), (size_t) 1, fp);
  fwrite( &aircost, sizeof(double), (size_t) 1, fp);
  fclose(fp);
  return row;

}

int getchyn_defaulty( void)
{
  int yn;
  do {
    yn=getch();
    if(yn==27 || yn==10 || yn==13 ) return 'Y';
    if(yn!='n' && yn!='N' && yn!='y' && yn!='Y') printf("%c",7);
  }while(yn!='y' && yn!='Y' && yn!='n' && yn!='N');
  return (yn | 0x20) ;
}

int getchyn( void)
{
  int yn;
  do {
    yn=getch();
    if(yn==27 || yn==10 || yn==13 ) return 'N';
    if(yn!='y' && yn!='Y' && yn!='n' && yn!='N') printf("%c",7);
  }while(yn!='y' && yn!='Y' && yn!='n' && yn!='N');
  return (yn | 0x20) ;
}

char *cgetsn( char *buffer, char *cn, char *default_string )
{
int i, c, j, finish=0, s;
struct rccoord rc, orgrc;
char strin[2];
char *t;

  c=cnumbuf[0];
  cnumbuf[1]=0;
  cnumbuf[2]=0;
  j=2;
  orgrc = rc = _gettextposition();
  if( vc.numcolors > 2) {
    _settextcolor(11);
  }
  _outtext(default_string);
  _settextposition( orgrc.row, orgrc.col);
  s = rc.col;
  if( vc.numcolors > 2) {
    _settextcolor(15);
  }
  do {
    strin[0] = (char) getch();
    if(j==2) {
      _settextposition( orgrc.row, orgrc.col);
       for(i=0; i<(c-1); i++) _outtext("_");
      _settextposition( orgrc.row, orgrc.col);
    }
    if(strin[0] == '\0') {
      strin[0] = (char) getch();
      switch( strin[0]) {
	case 72:  /* UP ARROW = non-delete backspace */
	  finish=1;
	  cnumbuf[2]=2;
	  cnumbuf[3]=0;
	  break;
	case 80:	/* DOWN ARROW = non-delete backspace */
	  finish=1;
	  cnumbuf[2]=1;
	  cnumbuf[3]=0;
	  break;
	case 75:  /* LEFT ARROW = non-delete backspace */
	  if( j > 2 ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
	case 83:  /* DEL = delete char at curpos */
	  if( j==(c+1) ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
      }     /* end of switch */
    } else {
      switch( strin[0]) {
	case 8:  /* BACKSPACE */
	  if( j > 2 ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
	case 13: /* CARRIAGE RETURN / LINEFEED */
	  cnumbuf[j]=0;
	  finish = 1;
	  break;
	case 'b':	/* b or B for bailout */
	case 'B':	/* b or B for bailout */
	case 'c':	/* c or C for Closed circuit */
	case 'C':	/* b or B for bailout */
	case 'F':	// Auto finish
	case 'f':	// Auto finish
	case 'a':	// Auto finish open circuit
	case 'A':	// Auto finish open circuit
	  strin[0] = strin[0] & 0x5F;
	  if( (cn[0]==strin[0] || cn[1]==strin[0] ) && ppo2) {
	    cnumbuf[2]=strin[0];
	    cnumbuf[3]=0;
	    finish = 1;
	    break;
	  }
	  if( (cn[0]==strin[0] || cn[1]==strin[0] ) ) {
	    cnumbuf[2]=strin[0];
	    cnumbuf[3]=0;
	    finish = 1;
	    break;
	  }
	  else {
	    printf("%c",7);
	  }
	  break;
	case 27:	/* ESC = abandon? */
	  cnumbuf[2]=0;
	  finish = 1;
	  break;
	default:
	  if( j<(c+1) ) {
	    if( (strin[0] > 47 && strin[0] < 58) || strin[0]==46 ) {
	      cnumbuf[j] = strin[0];
	      strin[1]=0;
	      _outtext( strin );
	      j++;
	    }
	  }
	  else {
	    printf("%c",7);
	  }
	  break;
      }     /* end of switch */
      if( j>(c+1) ) j--;
    }

  } while( !finish);
  if( vc.numcolors > 2)_settextcolor(7);
  return &cnumbuf[2];
}

char *cgetsa( char *buffer )
{
int c, j, finish=0, s;
struct rccoord rc;
unsigned char strin[2];
char *t;

  if( vc.numcolors > 2) _settextcolor(15);
  c=cnumbuf[0];
  cnumbuf[1]=0;
  cnumbuf[2]=0;
  j=2;
  rc = _gettextposition();
  s = rc.col;
  do {
    if( (strin[0] = (unsigned char) getch()) == '\0') {
      strin[0] = (unsigned char) getch();
      switch( strin[0]) {
	case 75:  /* LEFT ARROW = non-delete backspace */
	  if( j > 2 ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
	case 83:  /* DEL = delete char at curpos */
	  if( j==(c+1) ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
      }     /* end of switch */
    } else {
      switch( strin[0]) {
	case 8:  /* BACKSPACE */
	  if( j > 2 ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
	case 13: /* CARRIAGE RETURN / LINEFEED */
	  cnumbuf[j]=0;
	  finish = 1;
	  break;
	case 27:        /* ESC = abandon? */
	  cnumbuf[2]=0;
	  finish = 1;
	  break;
	default:
	  if( j<(c+1) ) {
	    if( (strin[0] > 31 && strin[0] < 127) || strin[0]==46 || strin[0]=='$' || strin[0]==156 ) {
	      cnumbuf[j] = strin[0];
	      strin[1]=0;
	      _outtext( strin );
	      j++;
	    }
	    else {
	      printf("%c",7);
	    }
	  }
	  else {
	    printf("%c",7);
	  }
	  break;
      }     /* end of switch */
      if( j>(c+1) ) j--;
    }

  } while( !finish);
  if( vc.numcolors > 2)_settextcolor(7);
  return &cnumbuf[2];
}

char *cgetsb( char *buffer )
{
int c, j, finish=0, s;
struct rccoord rc;
unsigned char strin[2];
char *t;

  if( vc.numcolors > 2) _settextcolor(15);
  c=cnumbuf[0];
  cnumbuf[1]=0;
  rc = _gettextposition();
  s = rc.col;
  _outtext(&cnumbuf[2]);

  j=2+strlen(&cnumbuf[2]);
  do {
    if( (strin[0] = (unsigned char) getch()) == '\0') {
      strin[0] = (unsigned char) getch();
      switch( strin[0]) {
	case 75:  /* LEFT ARROW = non-delete backspace */
	  if( j > 2 ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
	case 83:  /* DEL = delete char at curpos */
	  if( j==(c+1) ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
      }     /* end of switch */
    } else {
      switch( strin[0]) {
	case 8:  /* BACKSPACE */
	  if( j > 2 ) {
	    rc = _gettextposition();
	    _settextposition( rc.row, rc.col-1);
	    _outtext( " ");
	    _settextposition( rc.row, rc.col-1);
	    j--;
	  }
	  break;
	case 13: /* CARRIAGE RETURN / LINEFEED */
	  cnumbuf[j]=0;
	  finish = 1;
	  break;
	case 27:        /* ESC = abandon? */
	  cnumbuf[2]=0;
	  finish = 1;
	  break;
	default:
	  if( j<(c+1) ) {
	    if( (strin[0] > 31 && strin[0] < 127) || strin[0]==46 || strin[0]=='$' || strin[0]==156 ) {
	      cnumbuf[j] = strin[0];
	      strin[1]=0;
	      _outtext( strin );
	      j++;
	    }
	    else {
	      printf("%c",7);
	    }
	  }
	  else {
	    printf("%c",7);
	  }
	  break;
      }     /* end of switch */
      if( j>(c+1) ) j--;
    }

  } while( !finish);
  if( vc.numcolors > 2)_settextcolor(7);
  return &cnumbuf[2];
}

void savefillcosts(void)
{
    if( (fp=fopen("price.tot", "wb")) == NULL ) {
      exit(3);
    }
    rewind(fp);
    fwrite( &airfillpricetotal, sizeof(double), (size_t) 1, fp);
    fwrite( &o2fillpricetotal, sizeof(double), (size_t) 1, fp);
    fwrite( &hefillpricetotal, sizeof(double), (size_t) 1, fp);
    fwrite( &airfillcosttotal, sizeof(double), (size_t) 1, fp);
    fwrite( &o2fillcosttotal, sizeof(double), (size_t) 1, fp);
    fwrite( &hefillcosttotal, sizeof(double), (size_t) 1, fp);
    fclose(fp);
}

void graphicscreenprint( void)
{
  int bit, col, y, edge, i, j, x, xl, yl, xedge, yedge;
  short bk, gp;

  switch( _getbkcolor()) {
    case _BLACK:
      bk = 0;
      break;
    case _MODEFOFFTOON:
    case _BLUE:
      bk = 1;
      break;
    case _MODEFOFFTOHI:
    case _GREEN:
      bk = 2;
      break;
    case _MODEFONTOOFF:
    case _CYAN:
      bk = 3;
      break;
    case _MODEFON:
    case _RED:
      bk = 4;
      break;
    case _MODEFONTOHI:
    case _MAGENTA:
      bk = 5;
      break;
    case _MODEFHITOOFF:
    case _BROWN:
      bk = 6;
      break;
    case _MODEFHITOON:
    case _WHITE:
      bk = 7;
      break;
    case _MODEFHI:
    case _GRAY:
      bk = 8;
      break;
    case _LIGHTBLUE:
      bk = 9;
      break;
    case _LIGHTGREEN:
      bk = 10;
      break;
    case _LIGHTCYAN:
      bk = 11;
      break;
    case _LIGHTRED:
      bk = 12;
      break;
    case _LIGHTMAGENTA:
      bk = 13;
      break;
    case _YELLOW:
      bk = 14;
      break;
    case _BRIGHTWHITE:
      bk = 15;
      break;
    default:
      bk = 0;
      break;
  }
  open_output();

  if(lasermode==1) {
    /*
    _settextposition( 11,3);
    sprintf(stitle, "Line No ");
    _outtext( stitle);
    */
    edge = (int) vc.numxpixels;
    if( edge > 720) edge = 720;
    activate_lasergraphic_mode();
    laserleftmargin();
    for( y=0, xl=0, yl=300; y<vc.numypixels; y++, yl+=3) {
      laserposition( xl , yl );
      lasernumberofbytes(edge/8);
      for( x=0; x<edge; x +=8 ) {
	sprite[x] = 0;
	for( bit = 0; bit < 8; bit++) {
	  gp = _getpixel( x+bit , y);
	      if(vc.numcolors > 2) {
		  if( (gp != 7 && gp != 8) && !( (x+bit)>10 && (x+bit)<(edge-10) && y>46 && y<(vc.numypixels-301) ) ) { /* outside text window */
		    sprite[x] = sprite[x] + (short unsigned) pow( (double) 2, (double) (7-bit) );
		  }
		  if( !(gp != 7 && gp != 8) && ( (x+bit)>10 && (x+bit)<(edge-10) && y>46 && y<(vc.numypixels-301) ) ) { /* inside text window */
		    sprite[x] = sprite[x] + (short unsigned) pow( (double) 2, (double) (7-bit) );
		  }
	      }
	      else {
		  if( gp != bk ) {
		    sprite[x] = sprite[x] + (short unsigned) pow( (double) 2, (double) (7-bit) );
		  }
	      }
	}
	if( sprite[x] == 10) sprite[x] = 14;
	if( sprite[x] == 26) sprite[x] = 30;
	fwrite( &sprite[x], sizeof( char), (size_t) 1, fprn);
      }
    }
    deactivate_lasergraphic_mode();
    fprintf(fprn,"\n%c",12);
  }


  if(lasermode==2) {
    /*
    _settextposition( 11,3);
    sprintf(stitle, "Line No ");
    _outtext( stitle);
    */
    xedge = (int) vc.numxpixels;
    if( xedge > 720) xedge = 720;
    yedge = (int) vc.numypixels;
    if( yedge > 480) yedge = 480;
    activate_lasergraphic_mode75();
    laserleftmargin();
    for( y=0, xl=0, yl=150; y<xedge; y++, yl+=4) {
      laserposition( xl , yl );
      lasernumberofbytes(yedge/8);
      /*
      _settextposition( 11,11);
      sprintf(stitle, "%3d",y);
      _outtext( stitle);
      */
      for( x=yedge-1; x>0; x -=8 ) {
	sprite[x] = 0;
	for( bit = 0; bit < 8; bit++) {
	  gp = _getpixel( y, x-bit );
	      if(vc.numcolors > 2) {
		  if( (gp != 7 && gp != 8) && !( y>10 && y<(xedge-10) && (x-bit)>46 && (x-bit)<(vc.numypixels-301) ) ) { /* outside text window */
		    sprite[x] = sprite[x] + (short unsigned) pow( (double) 2, (double) (7-bit) );
		  }
		  if( !(gp != 7 && gp != 8) && ( y>10 && y<(xedge-10) && (x-bit)>46 && (x-bit)<(vc.numypixels-301) ) ) { /* outside text window */
		    sprite[x] = sprite[x] + (short unsigned) pow( (double) 2, (double) (7-bit) );
		  }
	      }
	      else {
		  if( gp != bk ) {
		    sprite[x] = sprite[x] + (short unsigned) pow( (double) 2, (double) (7-bit) );
		  }
	      }
	}
	if( sprite[x] == 10) sprite[x] = 14;
	if( sprite[x] == 26) sprite[x] = 30;
	fwrite( &sprite[x], sizeof( char), (size_t) 1, fprn);
      }
    }
    deactivate_lasergraphic_mode();
    fprintf(fprn,"\n%c",12);
  }

  if(!lasermode) {
    edge = (int) vc.numxpixels;
    if( edge > 800) edge = 800;
    edge -= (vc.numxpixels / vc.numtextcols)*2;
    for( col = (vc.numxpixels / vc.numtextcols)*2; col < edge; col += 8) {
      send_linefeed();
      activate_graphic_mode();
      for( y=479; y>=0; y--) {
	if(y<458 && y>10 && !(y<190 && y>159) ) {
	  sprite[y] = 0;
	  if( y < vc.numypixels) {
	    for( bit = 7; bit >= 0; bit--) {
	      gp = _getpixel( col+(7-bit), y);
	      if(vc.numcolors > 2) {
		if( y<160 && y>45 ) {
		  if( gp != 0 && gp != 8 && gp != 1 ) {
		    sprite[y] = sprite[y] + (short unsigned) pow( (double) 2, (double) bit);
		  }
		}
		else {
		  if( gp != 7 && gp != 8 ) {
		    sprite[y] = sprite[y] + (short unsigned) pow( (double) 2, (double) bit);
		  }
		}
	      }
	      else {
		  if( gp != bk ) {
		    sprite[y] = sprite[y] + (short unsigned) pow( (double) 2, (double) bit);
		  }
	      }
	    }
	  }
	  if( sprite[y] == 10) sprite[y] = 14;
	  if( sprite[y] == 26) sprite[y] = 30;
	  fwrite( &sprite[y], sizeof( char), (size_t) 1, fprn);
	  fwrite( &sprite[y], sizeof( char), (size_t) 1, fprn);
	}
      }
    }
    restore_linefeed();
    fprintf(fprn,"%c%c",27,64);
    fprintf(fprn,"\n\n");
  }
  fclose( fprn);
}

int getgas(int row, int col, double ppo2depth)
{
double ppo2temp;
int rowinit, colinit, redo, tablemix=-1;
int i;

  rowinit = row;
  colinit = col;
  nitrogenfraction = 0.001;
  heliumfraction = 0.001;
  ppo2fraction = 0.001;

  do {
    tablemix=-1;
    redo=0;
    if(air) {
      nitrogenfraction = 0.79;
      _settextposition( row, col);
      if(heo2display) sprintf(stitle, "O2=%d%%  ",(int)(OXYGENFRACTION*100.00+0.49) );
      else sprintf(stitle, "N2=%d%%  ",(int)(nitrogenfraction*100.00+0.49) );
      _outtext( stitle) ;
      nitrogenfractionlast = nitrogenfraction;
      return 0;
    }

    if(ppo2 && ppo2depth && !bailout) {
      helpscreen(6);
      ppo2fraction = ppo2fractionlast;
      heliumfractionlast = hemix= hemixlast ;
      nitrogenfractionlast = n2mix = n2mixlast;
      oxygenfractionlast = 1.00 - hemix - n2mix;
      if(autofinish) ;
      else {
	do {
	  _settextposition( row, col);
	  sprintf(stitle, "PPO2=____bar ");
	  _outtext( stitle) ;
	  ppo2print(ppo2fractionlast);
	  sprintf(stitle, "%1d.%02d", ppi, ppf);
	  _settextposition( row, col+5);
	  cnumbuf[0]=5;
	  numbuf = cgetsn( cnumbuf, "BA" ,stitle);
	  if(numbuf[0]==2) return -1;
	  ppo2fraction = ((double)atof( numbuf ));
	  if(!ppo2fraction ) {
	    ppo2fraction = ppo2fractionlast;
	    if(numbuf[0]=='B') bailout=1;
	    else bailout=bailoutlast;
	    if(numbuf[0]=='A') autofinish=1;
	    else autofinish=0;
	  }
	  else {
	    bailout=0;
	    autofinish=0;
	  }
	} while ( ( ppo2fraction > 2.00 || ppo2fraction <= 0.16 ) && !bailout );
	if(he && !bailout) {
	  if(automaticmode && ppo2depth) ;
	  else {
	    do {
	      helpscreen(50);
	      _settextposition( row, col);
	      sprintf(stitle, "He=__%%      ");
	      _outtext( stitle) ;
	      sprintf(stitle, "%2.0f", (hemixlast*100.00) );
	      _settextposition( row, col+3);
	      cnumbuf[0]=3;
	      numbuf = cgetsn( cnumbuf, "", stitle );
	      if(numbuf[0]==2) return -1;
	      if( (numbuf[0]>0x10) ) {
		hemix = ((double)atoi( numbuf ))/100.00;
	      }
	      else hemix = hemixlast;
	      if(depthdepth[0]==ppo2depth) {
		if((1.00-hemix-ppo2fraction/((ppo2depth/10.00) + atmospheric))<0.00) ;
		else n2mixlast=(1.00-hemix-ppo2fraction/((ppo2depth/10.00) + atmospheric));
	      }
	      helpscreen(51);
	      _settextposition( row, col);
	      sprintf(stitle, "O2=__%%      ");
	      _outtext( stitle) ;
	      if(n2mixlast && ((n2mixlast+hemix) <= 1.00) ) sprintf(stitle, "%2.0f%", ((1.00 - hemix - n2mixlast)*100.00) );
	      else sprintf(stitle, "%2.0f", ((1.00 - hemix)*100.00) );
	      _settextposition( row, col+3);
	      cnumbuf[0]=3;
	      numbuf = cgetsn( cnumbuf, "", stitle );
	      if(numbuf[0]==2) return -1;
	      if( (numbuf[0]>0x10) ) {
		n2mix = 1.00 - hemix - ((double)atoi( numbuf ))/100.00;
	      }
	      else {
		n2mix = n2mixlast;
	      }
	      if( (hemix+n2mix)>1.00 ) n2mix=0.00;
	      if( (n2mix<0.00) && (n2mix>(-0.0001)) ) n2mix=0.00; //Remove negative maths rounding error in hemix and n2mix divide by 100 above
	      ppo2temp = (ppo2depth/10.00 + atmospheric) * ( 1.00 - n2mix - hemix);
	      if( (ppo2temp<=0.16 && ppo2depth>=(stopfactor * 9.99) && ppo2depth) || ppo2temp>=1.80) {
		_settextposition( row, col);
		if(ppo2temp<=0.16) sprintf(stitle, " PPO2 LOW");
		else sprintf(stitle, " PPO2 HIGH");
		_outtext( stitle) ;
		delay1sec();
		_settextposition( row, col);
		sprintf(stitle, "           ");
		_outtext( stitle) ;
	      }

	    } while ( hemix<0.00 || hemix>1.00 || n2mix<0.00 || n2mix>1.00 || (bailout_breathable && ppo2temp<=0.16 && ppo2depth>=(stopfactor * 10.00) && ppo2depth ) || ppo2temp>=2.00 || (hemix==0.00 && n2mix==0.00) );
	    heliumfractionlast = hemixlast = hemix;
	    nitrogenfractionlast = n2mixlast = n2mix;
	    oxygenfractionlast = 1.00 - hemix - n2mix;
	  }
	  _settextposition( row, col);
	  sprintf(stitle, "He=%2.0f%% O2=%2.0f%%", hemix*100.00, (1.00-hemix-n2mix)*100.00 );
	  _outtext( stitle) ;
	  delay1sec();
	}
      }
      absolutedepth= (ppo2depth/10.00) + atmospheric ;
      absolutedepthpure= (ppo2depth/10.00) + atmospheric ;
      fractioncalcs(1); /*update ppo2fraction if current value can not be achieved at current depth*/
      _settextposition( row, col);
      ppo2print(ppo2fraction);
      ppo2fractionlast = ppo2fraction;
      bailoutlast = bailout;

      if(!bailout) {
	sprintf(stitle, "PPO2=%1d.%02dbar ", ppi, ppf);
	_outtext( stitle) ;
	col = col + 7;
	return 0;
      }
      _settextposition( row, col);
      sprintf(stitle, "Bailout      ");
      _outtext( stitle) ;
      delay1sec();
    }

    do {
      if(heo2display) {
	fractionmax = ( 0.161 / (ppo2depth/10.00 + atmospheric));
	tablemix=setoxygenfraction(ppo2depth); //Set for helpscreen, safe to do as oxygenfraction is updated properly later.
	if(bailout && ppo2depth) helpscreen(4);
	else if(ppo2depth) helpscreen(53);
	     else helpscreen(60);
	if(automaticmode && ppo2depth) ;
	else {
	  _settextposition( row, col);
	  sprintf(stitle, "O2=__%% ");
	  _outtext( stitle) ;
	  sprintf(stitle, "%d%",(int)(oxygenfraction*100.00+0.49) );
	  _settextposition( row, col+3);
	  cnumbuf[0]=3;
	  if(ppo2 && ppo2depth) {
	    numbuf = cgetsn( cnumbuf, "CA", stitle );
	    if(numbuf[0]==2) return -1;
	    if(numbuf[0]=='C') {
	      bailout=bailoutlast=0;
	      redo=1;
	      continue;
	    }
	    else bailout=bailoutlast;
	  }
	  else {
	    numbuf = cgetsn( cnumbuf, "A", stitle );
	    if(numbuf[0]==2) return -1;
	  }
	  if(numbuf[0]=='A') {
	    automaticmode=1;
	    //optimiseo2=1;
	  }
	  oxygenfraction = ((double)atoi( numbuf ))/100.00;
	  if(!oxygenfraction) {
	    tablemix=setoxygenfraction(ppo2depth);
	  }
	  else if(tablemix>=0) tablemix+=10;
	}//auto
	if( !oxygenfraction ) oxygenfraction = 0.21;
	if( oxygenfraction>0.99 ) oxygenfraction = 0.99;
	_settextposition( row, col);
	sprintf(stitle, "O2=%d%%  ",(int)(oxygenfraction*100.00+0.49) );
	_outtext( stitle) ;
	col = col + 7;
	//oxygenfractionlast = oxygenfraction;
	nitrogenfraction = 1.00 - oxygenfraction;
	if(he) {
	  if( (heliumfractionlast+oxygenfraction) >= 1.00 ) heliumfractionlast = 0.999 - oxygenfraction;
	  fractionmax = 1.00 - oxygenfraction ;
	  heliumfraction = heliumfractionlast;
	  if(automaticmode && ppo2depth) ;
	  else {
	    helpscreen(5);
	    _settextposition( row, col);
	    sprintf(stitle, "He=__%%");
	    _outtext( stitle) ;
	    sprintf(stitle, "%d%",(int)(heliumfractionlast*100.00+0.49) );
	    _settextposition( row, col+3);
	    cnumbuf[0]=3;
	    numbuf = cgetsn( cnumbuf, "", stitle );
	    if(numbuf[0]==2) return -1;
	    heliumfraction = ((double)atoi( numbuf ))/100.00 + 0.001;
	    if(!*numbuf ) heliumfraction = heliumfractionlast;
	    else if(tablemix>=0 && tablemix<10) tablemix+=10;
	  }
	  if( !heliumfraction ) heliumfraction = 0.001 ;
	  if( (heliumfraction+oxygenfraction) >= 1.00 ) heliumfraction = ONE_POINT - oxygenfraction;
	  _settextposition( row, col);
	  sprintf(stitle, "He=%d%% ",(int)(heliumfraction*100.00+0.49) );
	  _outtext( stitle) ;
	  col = col + 7;
	  //heliumfractionlast = heliumfraction;
	  nitrogenfraction = ONE_POINT - oxygenfraction - heliumfraction;
	}
	if(nitrogenfraction <= 0.00) nitrogenfraction=0.001;
	ppo2temp = (ppo2depth/10.00 + atmospheric) * ( 1.00 - nitrogenfraction - heliumfraction);
	if( ppo2temp <= 0.16 && ppo2depth) {
	  row = rowinit;
	  col = colinit;
	  _settextposition( row, col);
	  sprintf(stitle, " PPO2 TOO LOW");
	  _outtext( stitle) ;
	  delay1sec();
	  _settextposition( row, col);
	  sprintf(stitle, "             ");
	  _outtext( stitle) ;
	  automaticmode=autofinish=0;
	}
	if( ppo2temp > ppo2_limit_upper ) {
	  row = rowinit;
	  col = colinit;
	  _settextposition( row, col);
	  sprintf(stitle, "PPO2 TOO HIGH");
	  _outtext( stitle) ;
	  delay1sec();
	  _settextposition( row, col);
	  sprintf(stitle, "             ");
	  _outtext( stitle) ;
	  automaticmode=autofinish=0;
	}
      }
      else {
	fractionmax = 1.00 - ( 0.161 / (ppo2depth/10.00 + atmospheric));
	helpscreen(4);
	_settextposition( row, col);
	sprintf(stitle, "N2=__%% ");
	_outtext( stitle) ;
	_settextposition( row, col+3);
	cnumbuf[0]=3;
	numbuf = cgetsn( cnumbuf, "", "" );
	if(numbuf[0]==2) return -1;
	nitrogenfraction = ((double)atoi( numbuf ))/100.00;
	if(!*numbuf ) nitrogenfraction = nitrogenfractionlast;
	if( !nitrogenfraction ) nitrogenfraction = 0.001;
	_settextposition( row, col);
	sprintf(stitle, "N2=%d%%  ",(int)(nitrogenfraction*100.00+0.49) );
	_outtext( stitle) ;
	col = col + 7;
	//nitrogenfractionlast = nitrogenfraction;

	if(he) {
	  if( (heliumfractionlast+nitrogenfraction) >= 1.00 ) heliumfractionlast = 0.001 ;
	  fractionmax = 1.00 - ( 0.161 / (ppo2depth/10.00 + atmospheric)) - nitrogenfraction ;
	  helpscreen(5);
	  _settextposition( row, col);
	  sprintf(stitle, "He=__%%");
	  _outtext( stitle) ;
	  _settextposition( row, col+3);
	  cnumbuf[0]=3;
	  numbuf = cgetsn( cnumbuf, "", "" );
	  if(numbuf[0]==2) return -1;
	  heliumfraction = ((double)atoi( numbuf ))/100.00;
	  if(!*numbuf ) heliumfraction = heliumfractionlast;
	  if( !heliumfraction ) heliumfraction = 0.001 ;
	  if( (heliumfraction+nitrogenfraction) >= 1.00 ) heliumfraction = 0.001 ;
	  _settextposition( row, col);
	  sprintf(stitle, "He=%d%% ",(int)(heliumfraction*100.00+0.49) );
	  _outtext( stitle) ;
	  col = col + 7;
	  //heliumfractionlast = heliumfraction;
	}

	ppo2temp = (ppo2depth/10.00 + atmospheric) * ( 1.00 - nitrogenfraction - heliumfraction);
	if( ppo2temp <= 0.16 && ppo2depth) {
	  row = rowinit;
	  col = colinit;
	  _settextposition( row, col);
	  sprintf(stitle, " PPO2 TOO LOW");
	  _outtext( stitle) ;
	  delay1sec();
	  _settextposition( row, col);
	  sprintf(stitle, "             ");
	  _outtext( stitle) ;
	  automaticmode=autofinish=0;
	}
	if( ppo2temp > ppo2_limit_upper	) {
	  row = rowinit;
	  col = colinit;
	  _settextposition( row, col);
	  sprintf(stitle, "PPO2 TOO HIGH");
	  _outtext( stitle) ;
	  delay1sec();
	  _settextposition( row, col);
	  sprintf(stitle, "             ");
	  _outtext( stitle) ;
	  automaticmode=autofinish=0;
	}
      }
    } while( ((ppo2temp <= 0.16 && ppo2depth) || ppo2temp > ppo2_limit_upper) && !redo) ;
    if(tablemix>=10) {
      i=tablemix=tablemix%NUMGASMIX;
      if(gasused[i]==1 || gasstatus[i]==1 || i==0) {
	i++;
	i%=NUMGASMIX;
	for(;i!=tablemix;) {
	  if(i==0) i++;
	  if(i==tablemix) break;
	  if(gasstatus[i]==0) break;
	  i++;
	  i%=NUMGASMIX;
	}
	if(i==tablemix) { //Still no spare gas found
	  for(i=(NUMGASMIX-1);gasused[i] && i>0;i--);
	  if(i<1) i=tablemix; //use original gas
	}
      }
      tablemix=i;
      tablemix%=10;
      gasmixtable[tablemix][0] = nitrogenfraction;
      gasmixtable[tablemix][1] = heliumfraction;
      gasmixtable[tablemix][2] = ppo2depth;
    }
    if(set_gastable);
    else if(ppo2depth) {
      gasstatus[tablemix]=1;
      gasused[tablemix]=1;
    }
    oxygenfractionlast = oxygenfraction;
    heliumfractionlast = heliumfraction;
    nitrogenfractionlast = nitrogenfraction;
  } while(redo);
  return 0;
}

void settitletext(void)
{

  strcpy( options, "helv");
  strcat( strcat( strcpy( list, "t'"), options), "'");
  if(piyfact < 2) strcat( list, "h16w10b");
  else strcat( list, "h8w5b");
  _setfont( list);
  _getfontinfo( &fi);

}

void settitletextlarge(void)
{

  strcpy( options, "helv");
  strcat( strcat( strcpy( list, "t'"), options), "'");
  if(piyfact < 2) strcat( list, "h24w10b");
  else strcat( list, "h10w5b");
  _setfont( list);
  _getfontinfo( &fi);

}

void setaxistext(void)
{

  strcpy( options, "helv");
  strcat( strcat( strcpy( list, "t'"), options), "'");
  if(piyfact < 2) strcat( list, "h14w8b");
  else strcat( list, "h7w4b");
  _setfont( list);
  _getfontinfo( &fi);

}

void setmicroaxistext(void)
{

  strcpy( options, "helv");
  strcat( strcat( strcpy( list, "t'"), options), "'");
  if(piyfact < 2) strcat( list, "h12w6b");
  else strcat( list, "h6w3b");
  _setfont( list);
  _getfontinfo( &fi);

}

void titleprofilegraph(int xpos, double depthmax, double timetotald)
{
int k=0;
unsigned long j, timetotal;
unsigned char title[100];
       if(!strcmp(argvg,"bignose5a")) {
	 while(!kbhit()); getch();
       }

  if(xpos==xposgraph1) {  /*Clear axis scaling text*/
    if( vc.numcolors > 2)_setcolor(7);
    _rectangle( _GFILLINTERIOR, (x1graph1-30)/pixfact, (y1graph1-10)/piyfact, (x1graph1)/pixfact, (y2graph1+10)/piyfact );
    _rectangle( _GFILLINTERIOR, (x1graph1)/pixfact, (y2graph1)/piyfact, (x2graph1+9)/pixfact, (y2graph1+10)/piyfact );
    _rectangle( _GFILLINTERIOR, (x1graph1)/pixfact, (y2graph1)/piyfact, (x1graph1+170)/pixfact, (y2graph1+28)/piyfact );
  }

    timetotal = (long)timetotald;

    piyfactdec=piyfactdec/piyfact;
    setaxistext();
    if( vc.numcolors > 2)_setcolor(0);

    _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );

    _moveto( (25+xpos)/pixfact, 225/piyfact-piyfactdec);
    _outgtext( "D");
    _moveto( (25+xpos)/pixfact, 237/piyfact-piyfactdec);
    _outgtext( "e");
    _moveto( (25+xpos)/pixfact, 249/piyfact-piyfactdec);
    _outgtext( "p");
    _moveto( (25+xpos)/pixfact, 261/piyfact-piyfactdec);
    _outgtext( "t");
    _moveto( (25+xpos)/pixfact, 273/piyfact-piyfactdec);
    _outgtext( "h");


    _moveto( (25+xpos)/pixfact, 295/piyfact-piyfactdec);
    _outgtext( form);

    _moveto( (45+xpos)/pixfact, 215/piyfact-piyfactdec);
    _outgtext( "0");

    for(j=5; (depthmax * feetfactor) >= (j*3); j=j+5);

    _moveto( (35+xpos)/pixfact, 265/piyfact-piyfactdec);
    ultoa( j, title, 10);
    _outgtext( title);

    _moveto( (35+xpos)/pixfact, 315/piyfact-piyfactdec);
    ultoa( j*2, title, 10);
    _outgtext( title);

    _moveto( (35+xpos)/pixfact, 365/piyfact-piyfactdec);
    ultoa( j*3, title, 10);
    _outgtext( title);

  depthmaxgraph = 150.00 * feetfactor /(((double)j*3) * (double)piyfact);

    _moveto( (60+xpos)/pixfact, 372/piyfact-piyfactdec);
    _outgtext( "0");

    k=0;
    _moveto( (125+xpos)/pixfact, 384/piyfact-piyfactdec);
    if(timetotal>980) {
      k++;
      timetotal=timetotal/60;
      if(timetotal>980) {
	k++;
	_outgtext( "Time (days)");
	timetotal=timetotal/24;
      }
      else _outgtext( "Time (hours)");
    }
    else _outgtext( "Time (mins)");

    for(j=5; timetotal >= (j*4); j=j+5);
    _moveto( (112+xpos)/pixfact, 372/piyfact-piyfactdec);
    ultoa( j, title, 10);
    _outgtext( title);

    _moveto( (175+xpos)/pixfact, 372/piyfact-piyfactdec);
    ultoa( j*2, title, 10);
    _outgtext( title);

    _moveto( (238+xpos)/pixfact, 372/piyfact-piyfactdec);
    ultoa( j*3, title, 10);
    _outgtext( title);

    if(j*4>999) _moveto( (290+xpos)/pixfact, 372/piyfact-piyfactdec);
    else _moveto( (298+xpos)/pixfact, 372/piyfact-piyfactdec);
    ultoa( j*4, title, 10);
    _outgtext( title);

  timetotalgraph = 250.00/(((double)j*4) * (double)pixfact);
  if(k) {
    timetotalgraph = timetotalgraph / 60.00;
    k--;
    if(k) timetotalgraph = timetotalgraph / 24.00;
  }
  piyfactdec=piyfactdec*piyfact;

}

void borderdraw(void)
{
    _moveto_w( 0.00/(double)pixfact, 0.00);
      _lineto_w( 5.00/(double)pixfact, 00.00/(double)piyfact );
      _lineto_w( 5.00/(double)pixfact, 155.00/(double)piyfact );
    _moveto_w( 4.00/(double)pixfact, 0.00/(double)piyfact );
      _lineto_w( 4.00/(double)pixfact, 151.00/(double)piyfact );
      _lineto_w( 255.00/(double)pixfact, 151.00/(double)piyfact );
    _moveto_w( 0.00/(double)pixfact, 150.00/(double)piyfact );
      _lineto_w( 255.00/(double)pixfact, 150.00/(double)piyfact );
      _lineto_w( 255.00/(double)pixfact, 155.00/(double)piyfact );
    _moveto_w( 0.00/(double)pixfact, 50.00/(double)piyfact );
      _lineto_w( 5.00/(double)pixfact, 50.00/(double)piyfact );
    _moveto_w( 0.00/(double)pixfact, 100.00/(double)piyfact );
      _lineto_w( 5.00/(double)pixfact, 100.00/(double)piyfact );
    _moveto_w( 68.00/(double)pixfact, 150.00/(double)piyfact );
      _lineto_w( 68.00/(double)pixfact, 155.00/(double)piyfact );
    _moveto_w( 130.00/(double)pixfact, 150.00/(double)piyfact );
      _lineto_w( 130.00/(double)pixfact, 155.00/(double)piyfact );
    _moveto_w( 193.00/(double)pixfact, 150.00/(double)piyfact );
      _lineto_w( 193.00/(double)pixfact, 155.00/(double)piyfact );

}

void seto2optimise(void)
{
double sftemp;
char c;
int i;

  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SET PPO2 OPTIMIZATION");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("        LIMITS");
  if( vc.numcolors > 2) _setcolor(7);

  helpscreen(55);
  _settextposition( 12,10);
  sprintf(stitle,"Do you want to optimize PPO2 for open circuit dives? Y/<N> ");
  _outtext( stitle);
  c=getchyn();
  automaticmode=0;
  if( c=='y' || c=='Y' ) {
    optimiseo2=1;
  }
  else {
    optimiseo2=0;
    return;
  }
  sprintf(stitle,"%c",c);
  _outtext( stitle);

  /*
  helpscreen(57);
  _settextposition( 13,10);
  sprintf(stitle,"Do you want to automatically use default optimize values? Y/<N> ");
  _outtext( stitle);
  c=getchyn();
  if( c=='y' || c=='Y' ) {
    automaticmode=1;
  }
  else {
    automaticmode=0;
  }
  sprintf(stitle,"%c",c);
  _outtext( stitle);
  */

  helpscreen(55);
  do {
      _settextposition( 14,25);
      sprintf(stitle, "Current PPO2 lower limit= %3.2f Bar  ", ppo2_limit_lower );
      _outtext( stitle) ;
      _settextposition( 15,25);
      sprintf(stitle, "Enter PPO2 lower limit= ____ Bar  ");
      _outtext( stitle) ;
      _settextposition( 15,49);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      sftemp = (double)atof( numbuf ) ;
  } while( (sftemp > 1.60) || ( (sftemp <= 0.16) && *numbuf ) );
  if(*numbuf) ppo2_limit_lower = sftemp;

  helpscreen(56);
  do {
      _settextposition( 16,25);
      sprintf(stitle, "Current PPO2 upper limit= %3.2f Bar  ", ppo2_limit_upper );
      _outtext( stitle) ;
      _settextposition( 17,25);
      sprintf(stitle, "Enter PPO2 upper limit= ____ Bar  ");
      _outtext( stitle) ;
      _settextposition( 17,49);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      sftemp =(double)atof( numbuf ) ;
  } while( (sftemp > 1.60) || ( (sftemp < ppo2_limit_lower) && *numbuf ) );
  if(*numbuf) ppo2_limit_upper = sftemp;

}

int dispaydepthdatagas(void)
{
int i, j, k, m, ii;
unsigned char title[100], titlednum[10];
char c;
  set_gastable=1;
  he=n2=1;
  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("SET GAS TABLE VALUES");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("        LIMITS");
  if( vc.numcolors > 2) _setcolor(7);

  helpscreen(58);
  _settextposition( 12,10);
  sprintf(stitle,"Do you want to edit gas tables for Trimix and Nitrox dives? Y/<N> ");
  _outtext( stitle);
  c=getchyn();
  automaticmode=0;
  if( c=='y' || c=='Y' ) {
    gas_table=1;
  }
  else {
    gas_table=0;
    set_gastable=0;
    return 0;
  }
  _settextposition( 12,10);
  sprintf(stitle,"                                                                  ");
  _outtext( stitle);

  strcpy(title, divetitl);
  for (i=0,j=3,m=4; i<NUMGASMIX; i++) {
    if(i<0) i=0;
    j=3+(i%5)*15;
    m=(1+i/5)*6;
    helpscreen(62);
    depthdepth[i] = gasmixtable[i][2];
    maxdepthalarm=200.00;
    do {
      _settextposition( m,j);
      sprintf(stitle, "MODGas%d=___%s", i+1, form);
      _outtext( stitle) ;
      sprintf(stitle, "%d%",(int)(depthdepth[i] * feetfactor+0.49) );
      _settextposition( m,j+8);
      if(i==9)_settextposition( m,j+9);
      cnumbuf[0]=4;
      numbuf = cgetsn( cnumbuf, "A", stitle );
      if(numbuf[0]==2) {
	i-=2;
	break;
      }
      if(numbuf[0]=='A') {
	break;
      }
      depthdepth[i] = (double)atoi( numbuf ) / feetfactor;
      if(maxdepthalarm < depthdepth[i]) helpscreen(8);
    } while( (maxdepthalarm < depthdepth[i]) );
    if(numbuf[0]=='A') { break; }
    if(numbuf[0]==2) {
      _settextposition( m,j);
      sprintf(stitle, "                ");
      _outtext( stitle) ;
      _settextposition( m+1,j);
      _outtext( stitle) ;
      _settextposition( m+2,j);
      _outtext( stitle) ;
      _settextposition( m+3,j);
      _outtext( stitle) ;
      _settextposition( m+4,j);
      _outtext( stitle) ;
      continue;
    }

    if(depthdepth[i] || numbuf[0]==0x30) gasmixtable[i][2] = depthdepth[i];
    else depthdepth[i] = gasmixtable[i][2];
    if(numbuf[0]==0x30) continue;
    _settextposition( m,j);
    sprintf(stitle, "MODGas%d=%d%s  ", i+1, (int)(depthdepth[i] * feetfactor+0.49), form);
    stitle[16]=0;
    _outtext( stitle) ;

    if(getgas(m+1,j,depthdepth[i])<0) {
      i-=2;
      _settextposition( m,j);
      sprintf(stitle, "                ");
      _outtext( stitle) ;
      _settextposition( m+1,j);
      _outtext( stitle) ;
      _settextposition( m+2,j);
      _outtext( stitle) ;
      _settextposition( m+3,j);
      _outtext( stitle) ;
      _settextposition( m+4,j);
      _outtext( stitle) ;
      continue;
    }
    gasmixtable[i][0] = nitrogenfraction;
    gasmixtable[i][1] = heliumfraction;
    _settextposition( m+2, j);
    ppo2_now = ((depthdepth[i]/10.00) + atmospheric) * ( 1.00 - nitrogenfraction - heliumfraction);
    ppo2print(ppo2_now);
    sprintf(stitle, "PPO2=%1d.%02dbar  ", ppi, ppf);
    _outtext( stitle) ;
    helpscreen(63);
    do {
     _settextposition( m+3,j);
     if( vc.numcolors > 2) {
       _settextcolor(11);
     }
     if(gasstatus[i]) sprintf(stitle, "Active ");
     else sprintf(stitle, "Inhibit");
     _outtext( stitle) ;
     c=(char)getch();
     if(c=='\0') {
      c=(char)getch();
      if(c==72) {
	i-=2;
	_settextposition( m,j);
	sprintf(stitle, "                ");
	_outtext( stitle) ;
	_settextposition( m+1,j);
	_outtext( stitle) ;
	_settextposition( m+2,j);
	_outtext( stitle) ;
	_settextposition( m+3,j);
	_outtext( stitle) ;
	_settextposition( m+4,j);
	_outtext( stitle) ;
	break;
      }
      else continue;
     }
     if(c==' ') gasstatus[i]=!gasstatus[i];
     if(c=='I'||c=='i') gasstatus[i]=0;
     if(c=='A'||c=='a') gasstatus[i]=1;
    }while(c!=27 && c!=0x0d);

    if( vc.numcolors > 2) {
      _settextcolor(7);
    }
    _settextposition( m+3,j);
    _outtext( stitle) ;
    _settextposition( m+4,j);
    if(gasused[i]) sprintf(stitle, "Used ");
    else sprintf(stitle, "Not used");
    _outtext( stitle) ;
  }
  _settextposition( 18,10);
  sprintf(stitle,"Do you want to save gas tables for Trimix and Nitrox dives? Y/<N> ");
  _outtext( stitle);
  c=getchyn();
  if( c=='y' || c=='Y' ) {
    do {
      _settextposition( 4,30);
      sprintf(stitle, "Enter Dive name: ________");
      _outtext( stitle) ;
      _settextposition( 5,30);
      _outtext( "Press Backspace key to edit name") ;
      _settextposition( 4,47);
      cnumbuf[0]=9;
      cnumbuf[1]=9;
      cnumbuf[2]=0;
      strcat(cnumbuf, title);
      numbuf = cgetsb( cnumbuf );
      if(!*numbuf) return 0;
      strcpy(title, numbuf);
    } while( 0 );
    putdisc_gasmixdata(title);
  }
  set_gastable=0;
  return 0;
}

int setoxygenfraction(double ppo2depth)
{
double ppo2temp;
int i, bestdepth=-1;

  if( (optimiseo2 || automaticmode) && ppo2depth) {
    ppo2temp=oxygenfractionlast*(ppo2depth/10.00 + atmospheric);
    if(ppo2temp<ppo2_limit_lower || ppo2temp>ppo2_limit_upper) {
      oxygenfraction=ppo2_limit_upper/(ppo2depth/10.00 + atmospheric);
    }
    else {
      oxygenfraction = oxygenfractionlast;
    }

  }
  else oxygenfraction = oxygenfractionlast;

  if(gas_table && ppo2depth) {
    if(set_gastable) {
     bestdepth=NUMGASMIX-1;
      for(i=0; i<NUMGASMIX; i++) {
	if( (gasmixtable[i][2]-ppo2depth)>=0.00 && gasstatus[i]) {
	 if( gasmixtable[bestdepth][2] > (gasmixtable[i][2]) ) bestdepth=i;
	 else if( (gasmixtable[bestdepth][2]-ppo2depth)<0.00 )	bestdepth=i;
	}
      }
    }
    else {
     for(i=NUMGASMIX-1;i>=0 && !gasstatus[i];i--);
     bestdepth=i;
      for(i=0; i<NUMGASMIX; i++) {
       if( (gasmixtable[i][2]-ppo2depth)>=0.00 && gasstatus[i]) {
	if( gasmixtable[bestdepth][2] > (gasmixtable[i][2]) ) bestdepth=i;
	else if( (gasmixtable[bestdepth][2]-ppo2depth)<0.00 )  bestdepth=i;
       }
      }
     }
    nitrogenfractionlast = gasmixtable[bestdepth][0];
    heliumfractionlast = gasmixtable[bestdepth][1];
    oxygenfraction = ONE_POINT - heliumfractionlast - nitrogenfractionlast;
  }

  if( !oxygenfraction ) oxygenfraction = 0.21;
  if( oxygenfraction>0.99 ) oxygenfraction = 0.99;
  return bestdepth;
}

double ascenttime (double deepest_depth) {
  double timetoascend;
  if( (deepest_depth-ASCENTDEPTHFAST) > 0.00 ) timetoascend = (deepest_depth-ASCENTDEPTHFAST)/ASCENTRATEFAST + ASCENTTIMEMEDIUM;
  else if( (deepest_depth-ASCENTDEPTHMEDIUM) > 0.00 ) timetoascend = (deepest_depth-ASCENTDEPTHMEDIUM)/ASCENTRATEMEDIUM + ASCENTTIME;
       else timetoascend = (deepest_depth)/ASCENTRATE;

  return timetoascend;
}

double ascenttimediff (double deepest_depth, double depth) {
  double timetoascend;

  timetoascend = ascenttime(deepest_depth) - ascenttime(depth);
  return timetoascend;

}

void tissupdate(int stopsprocessed)
{
int i, ii, j;
	if(!strcmp(argvg,"bignosec")) {
	  printf("he%g", hemixpoint[divenumber][6] );
	}
  tissueorgtransfer();
  depthlast=0.00;
      for(ii=0; (ii<=5) && timepointb[divenumber][ii] && depthpoint[divenumber][ii]; ii++) {
	depth=(depthpoint[divenumber][ii] * 1.03) + depthinc;
	absolutedepthpure=depthpoint[divenumber][ii]/10.00 + atmospheric;
	exposuretime=timepointb[divenumber][ii];
	nitrogenfraction=nitrogenfractiondepth[ii]; //nitrogenpoint[divenumber][
	heliumfraction=heliumfractiondepth[ii];  //heliumpoint[divenumber][
	ppo2fraction=ppo2fractiondepth[ii];
	hemix = hemixpointfractiondepth[ii];
	n2mix = n2mixpointfractiondepth[ii];
	bailout = bailoutpointfractiondepth[ii];
	tissueupdate(0);
      }
  pambtolcalctrue();
  totaltimetosurface[divenumber] = (long) ( ascenttime(depthlast) +0.50 );
  absolutedepthpure = pambtolmindepth/10.00 + atmospheric;
  depth = pambtolmindepth;	     /* To allow for time to surface*/
  exposuretime = 0.01;
  tissueupdate(0);
  ii=6;
  //n2mixpoint[divenumber][ii]+0.0001;
  if(!strcmp(argvg,"bignosea")) printf("CCCCCCCCCCCCCCCC%d",stopsprocessed);
  for(j=numberstops ; j>(numberstops-stopsprocessed) ; ii++, j--) {
      //for(ii=0; ii<numberstops-stopsprocessed && timepointb[ii] && depthpoint[ii]; ii++) {
	depth=depthpoint[divenumber][ii];
	absolutedepthpure=depthpoint[divenumber][ii]/10.00 + atmospheric;
	exposuretime=timepointb[divenumber][ii]+timepointc[divenumber][ii];
	nitrogenfraction=nitrogenpoint[divenumber][ii];
	heliumfraction=heliumpoint[divenumber][ii];
	ppo2fraction=ppo2point[divenumber][ii];
	hemix = hemixpoint[divenumber][ii];
	n2mix = n2mixpoint[divenumber][ii];
	bailout = bailoutpoint[divenumber][ii];
	if(!strcmp(argvg,"bignosec")) {
	  printf("%d pp%g, ab%g, n2%g, he%g",ii, ppo2fraction, absolutedepthpure, n2mix, hemix );
	}

	tissueupdate(0);
	if(!strcmp(argvg,"bignose9")) {
	  printf("TISSUPf%d",j);
	}
	totaltimetosurface[divenumber] = totaltimetosurface[divenumber] + (long)(exposuretime + TIMEMOD); /*round up minutes*/
  }

}
