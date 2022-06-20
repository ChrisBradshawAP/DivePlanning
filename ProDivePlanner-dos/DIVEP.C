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

int yhe;
double sftemp=0.00, fillpressure=0.00, workingdepth=0.00, workingppo2=1.40, o2frac=0.00, n2frac=0.00, hefrac=0.00, narcpressure=4.00*0.79, hefill=0.00, airfill=0.00, o2fill=0.00, n2fill=0.00, current_fillpressure=0.00, currento2=0.00
  , currenthe=0.00, currentn2=0.00;
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
extern   short pixfact, piyfact, piyfactdec;
  /*unsigned char menunumber;*/
extern   unsigned char options[8];

extern   unsigned char safetytitlednum[10];
extern   unsigned char list[20];
extern   char fondir[_MAX_PATH];
extern   struct videoconfig vc;
extern   struct _fontinfo fi;
extern   short x, y, f;
extern   long prev_bk;
extern   int xint, yint;

extern char cnumbuf[MAXSTR];
extern char tmpbuf[MAXSTR];
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
extern double absolutedepth, absolutedepthpure, depth, atmospheric, exposuretime, flytolmins, pambtolmin, pambtolmindepth, pigttolstopminus1[16], nitrogenfraction, nitrogenfracdec[110], heliumfraction, heliumfracdec[110], microfracdec[110];
extern double ppo2fraction, ppo2fracdec[110], surftime, depthlast, currentstoppressure;
extern double nitrogenfractiondepth[10], exposuretimedepth[10], depthdepth[10], heliumfractiondepth[10], ppo2fractiondepth[10];
extern double hemixpointfractiondepth[10], n2mixpointfractiondepth[10];
extern double depthpoint[10][110], timepointa[10][110], timepointb[10][110], timepointc[10][110], nitrogenpoint[10][110], heliumpoint[10][110], ppo2point[10][110], totaltimepoint[10][110], hemixpoint[10][100], n2mixpoint[10][100];
extern double hemixpointfracdec[110], n2mixpointfracdec[110];
extern double depthmaxgraph, timetotalgraph, dailyotu[50], diveotu[10], missionotu, maxppo2[10], safetyfactor;
extern double micro_mode;
extern double feetfactor, stopfactor, cuft_ltr_factor, psifactor, maxdepthalarm, nitrogenfractioncalc, heliumfractioncalc;
/*double depthinc, decdepthinc, releaserate, ptolinc; Tom release */
extern double depthinc, decdepthinc, releaserate, ptolinc;
extern double nitrogenfractionlast, heliumfractionlast, fractionmax;
extern double gasmix[11][NUMGASMIX][3], gasmixbartime[10][NUMGASMIX], gasreservefraction[10][NUMGASMIX];
extern double gasmixtable[NUMGASMIX][3];
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
extern int heo2display, repeat;
extern double oxygenfraction, oxygenfractionlast;
extern double hemix, n2mix, hemixlast, n2mixlast;
extern int bailout_breathable;
extern int bailout, bailoutlast;
extern int bailoutpointfracdec[110], bailoutpointfractiondepth[10], bailoutpoint[10][100];
extern double ppo2_limit_lower, ppo2_limit_upper;
extern int optimiseo2, automaticmode, autofinish, gas_table;
extern double stoptimeplus[10][110];
extern char gasstatus[NUMGASMIX], gasused[NUMGASMIX], set_gastable;

extern void setinitial(void);
extern int tissueupdate(int temp_ascent);
extern void tissuetemptransfer(void);
extern void pambtolcalc(int temp_ascent);
extern void pambtolcalctrue(void);
extern double acalc(int i, int n2only, int tempcalc);
extern double bcalc(int i, int n2only, int tempcalc);
int getdecompressiontime(void);
extern void storedivedepthandtime();
int dispaydepthdata(void);
extern void depthplot(void);
void plotdepthdata(int fullrun, int stopsprocessed, int depthprocessed);
extern void decomcalc(void);
void currentcnsotudisplay(int x, int y);
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
extern int getgas(int,int,double);
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
void printtoprinter(int j);
void gasfillcalcs(
     int i,int j,int c,int k,int ii,int opxos, unsigned char title[100], double sftemp);
void gasfillsummary(
     int i,int j,int c,int k,int ii,int oppos, unsigned char title[100], double sftemp);
void numgas_calc(int j, int i);
void tradecylinderfill(void);
void tradegasprice(void);
void tradegascurrency(void);
double fillgetgas(void);
void helpscreen(int);
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
void savefillcosts(void);
void graphicscreenprint( void);
void seto2optimise(void);
int setoxygenfraction(double ppo2depth);
extern int dispaydepthdatagas(void);
extern void putdisc_gasmixdata(char *mix_file);
extern void getdisc_gasmixdata(char *mix_file);
extern void tissupdate(int stopsprocessed);
void showresults(int recalc);
void printmix(void);

extern FILE *fp, *fprn;

extern float stoplookup_factor[3][4];
void setcylinderfill(void)
{
int i, c, state;
char message[MAXSTR];
unsigned char title[100];

  drawbackground();
  _moveto( 100/pixfact, 11/piyfact);
  _outgtext("CYLINDER FILLING");
  _moveto( 100/pixfact, 26/piyfact);
  _outgtext("  CALCULATIONS");
  if( vc.numcolors > 2) _setcolor(7);
  state=0;
  do {
   if(state<0 || state>7) state==0;
   if(state==0) {
    do {
      helpscreen(11);
      _settextposition( 6,20);
      if( vc.numcolors > 2)_settextcolor(15);
      sprintf(stitle, "Enter Fill pressure= ____%s",porb);
      _outtext( stitle) ;
      sprintf(stitle, "%d",(int)(fillpressure*psifactor+0.49));
      _settextposition( 6,41);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", stitle );
      if(!*numbuf) break;
      sftemp = ( (double)atoi( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    if(*numbuf>10) fillpressure = sftemp/psifactor;
    _settextposition( 6,20);
    sprintf(stitle, "Fill pressure= %d%s           ",(int)(fillpressure*psifactor+0.49),porb);
    _outtext( stitle) ;
    if(workingppo2) showresults(1);
    state=1;
   }

   if(state==1) {
    do {
      helpscreen(12);
      _settextposition( 7,20);
      if( vc.numcolors > 2)_settextcolor(15);
      sprintf(stitle, "Enter Working depth= ___%s  ",form);
      _outtext( stitle) ;
      sprintf(stitle, "%d",(int)(workingdepth*feetfactor+0.49));
      _settextposition( 7,41);
      cnumbuf[0]=4;
      numbuf = cgetsn( cnumbuf, "", stitle );
      if(!*numbuf) break;
      if(numbuf[0]==2) {
	state-=2;
	break;
      }
      sftemp = ( (double)atoi( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    if(*numbuf>10) workingdepth = sftemp/feetfactor;
    _settextposition( 7,20);
    sprintf(stitle, "Working depth= %d%s          ",(int)(workingdepth*feetfactor+0.49),form);
    _outtext( stitle) ;
    if(workingppo2) showresults(1);
    state++;
   }

   if(state==2) {
    do {
      helpscreen(13);
      _settextposition( 8,20);
      if( vc.numcolors > 2)_settextcolor(15);
      sprintf(stitle, "Enter Working PPO2= ___ bar ");
      _outtext( stitle) ;
      ppo2print(workingppo2);
      sprintf(stitle, "%1d.%02d",ppi,ppf);
      _settextposition( 8,40);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", stitle );
      if(!*numbuf) break;
      if(numbuf[0]==2) {
	state-=2;
	break;
      }
      sftemp = ( (double)atof( numbuf ) ) ;
    } while( (sftemp > 2.00) || !*numbuf);
    if(*numbuf>10) workingppo2 = sftemp;
    _settextposition( 8,20);
    ppo2print(workingppo2);
    sprintf(stitle, "Working PPO2= %1d.%02dbar         ",ppi,ppf);
    _outtext( stitle) ;
    if(workingppo2) showresults(1);
    state++;
   }

   if(state==3) {
    do {
      helpscreen(14);
      _settextposition( 9,20);
      if( vc.numcolors > 2)_settextcolor(15);
      sprintf(stitle, "Enter personal air narcosis depth= ___%s  ",form);
      _outtext( stitle) ;
      sprintf(stitle, "%d",(int)(feetfactor*10.00*(narcpressure/0.79-atmospheric) +0.49));
      _settextposition( 9,55);
      cnumbuf[0]=4;
      numbuf = cgetsn( cnumbuf, "", stitle );
      if(!*numbuf) break;
      if(numbuf[0]==2) {
	state-=2;
	break;
      }
      sftemp = ( (double)atoi( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    if(*numbuf>10) narcpressure = (sftemp/feetfactor+10.00)*0.079;
    _settextposition( 9,20);
    sprintf(stitle, "Personal air narcosis depth= %d%s          ",(int)(feetfactor*10.00*(narcpressure/0.79-atmospheric) +0.49),form);
    _outtext( stitle) ;
    if(workingppo2) showresults(1);
    state++;
   }
   if(state==4) {
    if(workingdepth>50.00) yhe=1; else yhe=0;
    helpscreen(15);
    if( vc.numcolors > 2)_settextcolor(12);
    _settextposition( 10,20);
    if(!yhe) sprintf(stitle, "Is helium to be included in mix Y/<N> ");
    else sprintf(stitle, "Is helium to be included in mix <Y>/N ");
    _outtext( stitle) ;
    do {
      c=getch();
      if(c==27) return;
      if(c==13) break;
      else if(c!='y' && c!='Y' && c!='n' && c!='N') printf("%c",7);
    }while(c!='y' && c!='Y' && c!='n' && c!='N');
    if( vc.numcolors > 2)_settextcolor(7);
    if(c=='y' || c=='Y') {
      yhe=1;
    }
    if(c=='n' || c=='N') {
      yhe=0;
    }
    if(yhe)	sprintf(stitle, "Y");
    else	sprintf(stitle, "N");
    _outtext( stitle) ;
    if(workingppo2) showresults(1);
    state++;
   }
    _settextposition( 12,20);
    if( vc.numcolors > 2)_settextcolor(7);
    sprintf(stitle, "%.1f%%O2, %.1f%%He, %.1f%%N2", (o2frac*100.00), (hefrac*100.00), (n2frac*100.00) );
    _outtext( stitle) ;
    if(state==5) {
      do {
	helpscreen(35);
	if( vc.numcolors > 2)_settextcolor(15);
	sprintf(stitle, "%.1f",(o2frac*100.00));
	_settextposition(12,20);
	cnumbuf[0]=5;
	numbuf = cgetsn( cnumbuf, "", stitle );
	if(!*numbuf) break;
	if(numbuf[0]==2) {
	  state-=3;
	  break;
	}
	sftemp = ( (double)atoi( numbuf ) ) ;
      } while( (sftemp < 0.00) || !*numbuf);
      if(*numbuf>10) o2frac = sftemp/100.00;
      _settextposition(12,20);
      if(workingppo2) showresults(0);
      state++;
    }
    if(state==6) {
      do {
	helpscreen(34);
	if( vc.numcolors > 2)_settextcolor(15);
	sprintf(stitle, "%.1f",(hefrac*100.00));
	_settextposition(12,29);
	cnumbuf[0]=5;
	numbuf = cgetsn( cnumbuf, "", stitle );
	if(!*numbuf) break;
	if(numbuf[0]==2) {
	  state-=2;
	  break;
	}
	sftemp = ( (double)atoi( numbuf ) ) ;
      } while( (sftemp < 0.00) || !*numbuf);
      if(*numbuf>10) hefrac = sftemp/100.00;
      _settextposition(12,20);
      if(workingppo2) showresults(0);
      state++;
    }

    _settextposition( 13,20);
    sprintf(stitle, "Mix narcosis depth= %d%s",(int)(feetfactor*10.00*(narcpressure/n2frac-atmospheric) +0.49),form);
    _outtext( stitle) ;
   if(state==7) {
    do {
      helpscreen(11);
      _settextposition(14,20);
      if( vc.numcolors > 2)_settextcolor(15);
      sprintf(stitle, "Enter CURRENT Fill pressure= ____%s",porb);
      _outtext( stitle) ;
      sprintf(stitle, "%d",(int)(current_fillpressure*psifactor+0.49));
      _settextposition(14,48);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", stitle );
      if(!*numbuf) break;
      if(numbuf[0]==2) {
	state-=3;
	break;
      }
      sftemp = ( (double)atoi( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    if(*numbuf>10) current_fillpressure = sftemp/psifactor;
    _settextposition(14,20);
    sprintf(stitle, "CURRENT Fill pressure= %d%s           ",(int)(current_fillpressure*psifactor+0.49),porb);
    _outtext( stitle) ;
    if(workingppo2) showresults(0);
    state++;
   }

   if(state==8) {
    if(current_fillpressure) {
      do {
	helpscreen(35);
	_settextposition(15,20);
	if( vc.numcolors > 2)_settextcolor(15);
	sprintf(stitle, "Enter Current O2 __%%");
	_outtext( stitle) ;
	sprintf(stitle, "%d",(int)(currento2*100.00));
	_settextposition(15,37);
	cnumbuf[0]=3;
	numbuf = cgetsn( cnumbuf, "", stitle );
	if(!*numbuf) break;
	if(numbuf[0]==2) {
	  state-=2;
	  break;
	}
	sftemp = ( (double)atoi( numbuf ) ) ;
      } while( (sftemp < 0.00) || !*numbuf);
      if(*numbuf>10) currento2 = sftemp/100.00;
      _settextposition(15,20);
      sprintf(stitle, "Current O2 = %d%%           ",(int)(currento2*100.00));
      _outtext( stitle) ;
    }
    if(workingppo2) showresults(0);
    state++;
   }

   if(state==9) {
    if(current_fillpressure) {
      do {
	helpscreen(34);
	_settextposition(16,20);
	if( vc.numcolors > 2)_settextcolor(15);
	sprintf(stitle, "Enter Current He __%%");
	_outtext( stitle) ;
	sprintf(stitle, "%d",(int)(currenthe*100.00));
	_settextposition(16,37);
	cnumbuf[0]=3;
	numbuf = cgetsn( cnumbuf, "", stitle );
	if(!*numbuf) break;
	if(numbuf[0]==2) {
	  state-=2;
	  break;
	}
	sftemp = ( (double)atoi( numbuf ) ) ;
      } while( (sftemp < 0.00) || !*numbuf);
      if(*numbuf>10) currenthe = sftemp/100.00;
      _settextposition(16,20);
      sprintf(stitle, "Current He = %d%%           ",(int)(currenthe*100.00));
      _outtext( stitle) ;
    }
    if(workingppo2) showresults(0);
    state++;
   }
   if(state>9) {
    helpscreen(100);
    _settextposition( 22,20);
    if( vc.numcolors > 2)_settextcolor(15);
    sprintf(stitle, "Re-Edit? Y/<N> ");
     _outtext( stitle) ;

     c = getchyn();
     if(c=='y' || c=='Y') state=0;
     else state=20;
   }
   _settextposition( 22,20);
   sprintf(stitle, "                  ");
   _outtext( stitle) ;
  } while(state<20);
    if( vc.numcolors > 2)_settextcolor(7);
    _settextposition( 22,20);
    sprintf(stitle, "Send to printer? Y/<N> ");
    _outtext( stitle) ;

    c = getchyn();
    if(c=='y' || c=='Y') {

      if((fprn=fopen( "PRN", "w"))==NULL) {
	strcpy( message, "Unable to open printer file.\n");
	write(fileno(stdout), message, strlen(message));
	exit(3);
      }
      printmix();
      fprintf(fprn,"\n\n%c",12);
      fclose( fprn);
    }
    _settextposition( 22,20);
    sprintf(stitle, "Append to file gasfill.txt? Y/<N> ");
    _outtext( stitle) ;

    c = getchyn();
    if(c=='y' || c=='Y') {

      if((fprn=fopen( "gasfill.txt", "a+"))==NULL) {
	strcpy( message, "Unable to open printer file.\n");
	write(fileno(stdout), message, strlen(message));
	exit(3);
      }
      printmix();
      fclose( fprn);
    }
}

void showresults(int recalc)
{
    if( vc.numcolors > 2) _setcolor(7);
  if(recalc) {
    o2frac = workingppo2 / ((workingdepth/10.00) + atmospheric);
    n2frac = 1.00 - o2frac;
    hefrac = 0.00;
    if(yhe) {
      if( (n2frac * (workingdepth/10.00 + atmospheric)) > narcpressure )
	n2frac = narcpressure / (workingdepth/10.00 + atmospheric);
    hefrac = 1.00 - n2frac - o2frac;
    }
  }
  else {
    workingppo2 = o2frac * ((workingdepth/10.00) + atmospheric);
    n2frac = 1.00 - o2frac - hefrac;
    narcpressure = n2frac * (workingdepth/10.00 + atmospheric);
  }
    _settextposition( 6,20);
    sprintf(stitle, "Fill pressure= %d%s           ",(int)(fillpressure*psifactor+0.49),porb);
    _outtext( stitle) ;
    _settextposition( 7,20);
    sprintf(stitle, "Working depth= %d%s          ",(int)(workingdepth*feetfactor+0.49),form);
    _outtext( stitle) ;
    _settextposition( 8,20);
    ppo2print(workingppo2);
    sprintf(stitle, "Working PPO2= %1d.%02dbar         ",ppi,ppf);
    _outtext( stitle) ;
    _settextposition( 9,20);
    sprintf(stitle, "Personal air narcosis depth= %d%s          ",(int)(feetfactor*10.00*(narcpressure/0.79-atmospheric) +0.49),form);
    _outtext( stitle) ;
    _settextposition( 10,20);
    if(yhe)	sprintf(stitle, "Helium = Y                                 ");
    else	sprintf(stitle, "Helium = N                                 ");
    _outtext( stitle) ;
    _settextposition( 12,20);
    sprintf(stitle, "%.1f%%O2, %.1f%%He, %.1f%%N2", (o2frac*100.00), (hefrac*100.00), (n2frac*100.00) );
    _outtext( stitle) ;
    _settextposition( 13,20);
    sprintf(stitle, "Mix narcosis depth= %d%s",(int)(feetfactor*10.00*(narcpressure/n2frac-atmospheric) +0.49),form);
    _outtext( stitle) ;
    _settextposition(14,20);
    sprintf(stitle, "CURRENT Fill pressure= %d%s           ",(int)(current_fillpressure*psifactor+0.49),porb);
    _outtext( stitle) ;
    _settextposition(15,20);
    sprintf(stitle, "Current O2 = %d%%           ",(int)(currento2*100.00+0.49));
    _outtext( stitle) ;
    _settextposition(16,20);
    sprintf(stitle, "Current He = %d%%           ",(int)(currenthe*100.00+0.49));
    _outtext( stitle) ;
    currentn2 = 1.00 - currento2 - currenthe;
    hefill = (hefrac * fillpressure) - (currenthe * current_fillpressure);
    airfill = ( n2frac * fillpressure - (currentn2 * current_fillpressure) ) / 0.79;
    o2fill = fillpressure - hefill - airfill - (current_fillpressure);
    _settextposition(18,20);
    sprintf(stitle, "Helium fill=%.1f%s    ",hefill*psifactor,porb);
    _outtext( stitle) ;
    _settextposition( 19,20);
    sprintf(stitle, "Air fill=%.1f%s    ",airfill*psifactor,porb);
    _outtext( stitle) ;
    _settextposition(20,20);
    sprintf(stitle, "Oxygen fill=%.1f%s    ",o2fill*psifactor,porb);
    _outtext( stitle) ;
}

void printmix(void)
{
      fprintf(fprn,    "\n\n********************************************************************************");
      fprintf(fprn,"\n          Fill pressure= %d%s	   ",(int)(fillpressure*psifactor+0.49),porb);
      fprintf(fprn,"\n          Working depth= %d%s          ",(int)(workingdepth*feetfactor+0.49),form);
      fprintf(fprn,"\n          Working PPO2= %1d.%02dbar         ",ppi,ppf);
      fprintf(fprn,"\n          %.1f%%O2, %.1f%%He, %.1f%%N2", (o2frac*100.00), (hefrac*100.00), (n2frac*100.00) );
      fprintf(fprn,"\n          Current mix narcosis depth= %d%s",(int)(feetfactor*10.00*(narcpressure/n2frac-atmospheric) +0.49),form);
      if(current_fillpressure) {
	fprintf(fprn,"\n          CURRENT Fill pressure= %d%s",(int)(current_fillpressure*psifactor+0.49),porb);
	fprintf(fprn,"\n          Current O2%=%d%%",(int)(currento2*100.00));
	fprintf(fprn,"\n          Current He%=%d%%",(int)(currenthe*100.00));
      }
      fprintf(fprn,"\n");
      if(hefill) fprintf(fprn,"\n          Helium fill=%.1f%s",hefill*psifactor,porb);
      if(airfill) fprintf(fprn,"\n          Air fill=%.1f%s",airfill*psifactor,porb);
      if(o2fill) fprintf(fprn,"\n          Oxygen fill=%.1f%s",o2fill*psifactor,porb);
      if(n2fill) fprintf(fprn,"\n          Nitrogen fill=%.1f%s",n2fill*psifactor,porb);
      _strdate( tradedatebuf );
      if(feetfactor=1.00) fprintf(fprn, "\n\n          Date:%c%c/%c%c/%c%c", tradedatebuf[3], tradedatebuf[4], tradedatebuf[0], tradedatebuf[1], tradedatebuf[6], tradedatebuf[7]);
      else	       fprintf(fprn, "\n\n          Date:%c%c/%c%c/%c%c", tradedatebuf[0], tradedatebuf[1], tradedatebuf[3], tradedatebuf[4], tradedatebuf[6], tradedatebuf[7]);
      fprintf(fprn,    "\n\n********************************************************************************");

}

void helpscreen(int lineno)
{
unsigned char title[100], titlea[100], titlec[100];
  titlec[0]=0;
  _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );
  setaxistext();
  x=310/pixfact; y=29/piyfact;
  if(lineno==41 || lineno==42 || lineno==52 || lineno==26) { x=20/pixfact; y=14; }
  if(lineno==4 && bailout) { x=300/pixfact; }
  if(lineno==6 && !bailout) { x=300/pixfact; }
  if( vc.numcolors > 2) {
    _setcolor(7);
    if( !(lineno==41 || lineno==42 || lineno==52 || lineno==26) ) _rectangle( _GFILLINTERIOR, vc.numxpixels-10/pixfact, 44/piyfact, 300/pixfact, y ); //clear helpscreen area
    else _rectangle( _GFILLINTERIOR, vc.numxpixels-10/pixfact, 44/piyfact, x, y ); //clear helpscreen area
    _setcolor(6);
  }
  else {
    _setviewport( x, y, 640/pixfact,44/piyfact );
    _clearscreen( _GVIEWPORT);
    _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );
  }
  strcpy(titlea,"HELP:");
  _moveto( (x+5/pixfact), y);

    switch(lineno) {

      case 0:
	sprintf(title, "              ");
	break;

      case 1:
	sprintf(title, " Enter Depth in %s", formlong);
	sprintf(titlec, " Commands: 0 = Abort, return= default");
	break;

      case 2:
	sprintf(title, " Enter Depth, 0 = finish");
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 3:
	sprintf(title, " Enter Time in minutes");
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 4:
	if(heo2display) {
	  if(ppo2 && bailout) {
	    sprintf(title, "Enter O2%% min=%d%%",(int)(fractionmax*100.00+1.00) );// ,(int)(oxygenfraction*100.00+0.49) );
	    sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default, A = Autofinish, C = Closed circuit");
	  }
	  else {
	    sprintf(title, " Enter O2 %%(min=%d%%)",(int)(fractionmax*100.00+1.00) );// ,(int)(oxygenfraction*100.00+0.49) );
	    sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default, A = Autofinish");
	  }
	}
	else {
	  sprintf(title, " Enter N2 %%(max=%d%%), return=%d%%",(int)(fractionmax*100.00) ,(int)(nitrogenfractionlast*100.00+0.49) );
	}
	break;

      case 5:
	sprintf(title, " Enter He %%(max=%d%%)",(int)(fractionmax*100.00+0.49) );// ,(int)(heliumfractionlast*100.00+0.49) );
	  sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 6:
	ppo2print(ppo2fractionlast);
	if(!bailout) {
	  sprintf(title, " Enter PPO2 " );//, ppi, ppf);
	  sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default, B = Bailout, A = Autofinish");
	}
	else {
	  sprintf(title, " Enter PPO2 bar, return=Bailout");
	  sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	}
	break;

      case 7:
	sprintf(title, " Enter Surface time minutes, 0=exit");
	break;

      case 8:
	if( vc.numcolors > 2)_setcolor(4);
	sprintf(title, " Enter new depth, maximum=%3d%s",(int)(maxdepthalarm*feetfactor+0.49),form);
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 9:
	if( vc.numcolors > 2)_setcolor(4);
	sprintf(title, " Enter new depth, minimum=%3d%s",(int)(pambtolmindepth*feetfactor+1.00),form);
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 10:
	sprintf(title, " More deco-stops, press any key....");
	break;

      case 11:
	sprintf(title, " Enter cylinder fill pressure");
	break;

      case 12:
	sprintf(title, " Enter maximum working depth");
	break;

      case 13:
	sprintf(title, " Enter maximum working PPO2");
	break;

      case 14:
	sprintf(title, " Depth you feel narcosis on air dive");
	break;

      case 15:
	sprintf(title, " Enter y if helium is to be used");
	break;

      case 16:
	sprintf(title, " Enter y for hard copy");
	break;

      case 17:
	sprintf(title, " Enter file name, or return= abort");
	break;

      case 18:
	sprintf(title, " Enter file name, return= current dive");
	break;

      case 19:
	sprintf(title, " Enter safety factor 0% to 50%");
	break;

      case 20:
	sprintf(title, " Enter atmospheric pressure in mbar");
	break;

      case 21:
	sprintf(title, " Enter hour, then press return");
	break;

      case 22:
	sprintf(title, " Enter minutes, then press return");
	break;

      case 23:
	sprintf(title, " Enter surface breathing rate");
	break;

      case 24:
	if( vc.numcolors > 2)_setcolor(4);
	sprintf(title, " Enter new rate, maximum=%3d%s",(int)(maxbreathratenumber),cuftorltrmin);
	break;

      case 25:
	if( vc.numcolors > 2)_setcolor(4);
	sprintf(title, " Enter new rate, minimum=%3d%s",(int)(minbreathratenumber),cuftorltrmin);
	break;

      case 26:
	if( vc.numcolors > 2)_setcolor(4);
	sprintf(title, " Enter y if mission to be continued");
	break;

      case 27:
	sprintf(title, " Enter gas %% in final mix          ");
	break;

      case 28:
	sprintf(title, " Oxygen %% too high with air top off");
	break;

      case 29:
	sprintf(title, " Enter n if cylinder empty         ");
	break;

      case 30:
	sprintf(title, " Enter y if helium is to be used");
	break;

      case 31:
	sprintf(title, " Enter P for %%, R for pressure ");
	break;

      case 32:
	sprintf(title, " Enter He pressure in cylinder  ");
	break;

      case 33:
	sprintf(title, " Enter O2 pressure in cylinder  ");
	break;

      case 34:
	sprintf(title, " Enter He percentage in cylinder");
	break;

      case 35:
	sprintf(title, " Enter O2 percentage in cylinder");
	break;

      case 36:
	sprintf(title, " Enter current cylinder pressure");
	break;

      case 37:
	sprintf(title, " Enter required cylinder pressure");
	break;

      case 38:
	sprintf(title, " Enter required cylinder pressure");
	break;

      case 39:
	sprintf(title, " Enter y for hard copy           ");
	break;

      case 40:
	sprintf(title, " Pressure too low, minimum=%.0f%s",currentfillpressure*psifactor,porb);
	break;

      case 41:
	sprintf(title, " Enter ammount in %s\\%s, return=no change", gascurrency, cuftorltrnos);
	break;

      case 42:
	sprintf(title, " Enter currency, eg:$, return=no change");
	break;

      case 43:
	sprintf(title, " Oxygen %% too low with air top off");
	break;

      case 44:
	sprintf(title, " Enter maximum cylinder pressure");
	break;

      case 45:
	sprintf(title, " Enter water capacity of cylinder");
	break;

      case 46:
	sprintf(title, " Enter free air capacity of cylinder");
	break;

      case 47:
	sprintf(title, " Pressure too high max=%.0f%s", trademaxcylinderpressure*psifactor,porb);
	break;

      case 48:
	sprintf(title, " Enter divers name, no name=END");
	break;

      case 49:
	sprintf(title, " Enter y to save to file gasfill.txt");
	break;

      case 50:
	sprintf(title, " Enter Diluent He %%", (hemixlast*100.00) );
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 51:
	  if(n2mixlast && ((n2mixlast+hemix) <= 1.00) ) sprintf(title, " Enter Diluent O2 %% );//, return=%3.0f%%", ((1.00 - hemix - n2mixlast)*100.00) );
	  else sprintf(title, " Enter Diluent O2 %%, return=%3.0f%%", ((1.00 - hemix)*100.00) );
	  sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 52:
	sprintf(title, " Enter Y to check oxygen content of Diluent is suitable for breathing at depth" );
	break;

      case 53:
	if(heo2display) {
	  sprintf(title, " Enter O2 %%(min=%d%%)",(int)(fractionmax*100.00+1.00) );// ,(int)(oxygenfraction*100.00+0.49) );
	}
	else sprintf(title, " Enter N2 %%(max=%d%%), return=%d%%",(int)(fractionmax*100.00) ,(int)(nitrogenfractionlast*100.00+0.49) );
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default, A = Autofinish");
	break;

      case 54:
	sprintf(title, " Enter Lower limit, ret=no change" );
	break;

      case 55:
	sprintf(title, " Enter Y for PPO2 optimization" );
	break;

      case 56:
	sprintf(title, " Enter Upper limit, ret=no change" );
	break;

      case 57:
	sprintf(title, " Enter Y for Automatic optimization" );
	break;

      case 58:
	sprintf(title, " Enter Y for Gas Table details" );
	break;

      case 59:
	sprintf(title, " Enter Y to re-edit dive" );
	break;

      case 60:
	if(heo2display) {
	  sprintf(title, " Enter O2 %%(min=%d%%)",(int)(fractionmax*100.00+1.00) );// ,(int)(oxygenfraction*100.00+0.49) );
	}
	else sprintf(title, " Enter N2 %%(max=%d%%), return=%d%%",(int)(fractionmax*100.00) ,(int)(nitrogenfractionlast*100.00+0.49) );
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 61:
	sprintf(title, " Enter extra stop time, if required" );
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default");
	break;

      case 62:
	sprintf(title, " Enter Depth");
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default,  A = Autofinish,  0 = Ignore");
	break;

      case 63:
	sprintf(title, " Select Gas Active/Inactive");
	sprintf(titlec, " Commands: Up arrow = Backedit,  Return = Default,  A = Active,  I = Inactive");
	break;

      default:
	sprintf(title, " Help line");


    }
  strcat(titlea,title);
  _outgtext(titlea);
  if( vc.numcolors > 2) {
    _setcolor(7);
    _rectangle( _GFILLINTERIOR, 6/pixfact, (vc.numypixels-19)/piyfact, vc.numxpixels-6, (vc.numypixels-6)/piyfact); //clear helpscreen area
    _setcolor(0);
  }
  if(titlec[0]) {
    setaxistext();
    _moveto( 5/pixfact, (vc.numypixels-20)/piyfact );
    _outgtext( titlec);
  }


}

int getdecompressiontime(void)
{
int i, j, p, k, m, n, oxtoolow=0, cont, tablemix;
double stopdepthtemp, maxdepth, micro_stop_depth[NUM_MICRO_STOPS], micro_stop_time[NUM_MICRO_STOPS];

  for(i=0;i<110;i++) {
    microfracdec[i]=0.00;
  }
  pambtolcalctrue();
  totaltimetosurface[divenumber] = (long) ( ascenttime(depthlast) +0.50 );
  absolutedepthpure = pambtolmindepth/10.00 + atmospheric;
  depth = pambtolmindepth;	     /* To allow for time to surface*/
  exposuretime = 0.01;
  tissueupdate(0);
  currentcnsotudisplay(11, 3);

  stopdepthtemp = (pambtolmin - atmospheric)*10.00;
    for(i=0,j=0; (stopdepthtemp - (double)i ) > 0.00; i=i+3, j++) ;
  deepeststop = i;
  numberstops = j;
  if(numberstops==1 && !stoplookup_factor[sixstopmode][j]) numberstops++;
  timetofirststop[divenumber] = totaltimetosurface[divenumber] - (long) ascenttime(deepeststop);
  if(numberstops>100) {
    _settextposition( 4,46);
    sprintf(stitle, "WARNING: OXYGEN TOO LOW");
    _outtext( stitle) ;
    _settextposition( 5,46);
    sprintf(stitle, "DIVE ABORTED: press any key");
    _outtext( stitle) ;
    getch();
    return 0;
  }
  maxdepth=0.00;
  for(i=0;i<6;i++) {
    if(depthpoint[divenumber][i]>maxdepth && timepointb[divenumber][i]) maxdepth=depthpoint[divenumber][i];
  }
    if(maxdepth) { //>50.00) {
     if(numberstops && micro_mode) {
       for(i=0;i<NUM_MICRO_STOPS;i++) {
	   micro_stop_depth[i]=0.00;
	   micro_stop_time[i]=0.00;
       }
       if( (maxdepth-stopdepthtemp)>20.00) {
	 micro_stop_depth[0]=((maxdepth-stopdepthtemp)/2.00+stopdepthtemp);
	 micro_stop_time[0]=1.00;
       }
       for(i=1;i<NUM_MICRO_STOPS;i++) {
	 if( (micro_stop_depth[i-1]) > (stopdepthtemp+MICRO_STEP) ) {
	   micro_stop_depth[i]=(micro_stop_depth[i-1]-stopdepthtemp)/2.00 + stopdepthtemp;
	   micro_stop_time[i]=1.00;
	 }
       }
       if(micro_stop_depth[0]) {
	for(i=NUM_MICRO_STOPS-1;i>=0;i--) {
	  if(micro_stop_depth[i]) {
	    numberstops++;
	    microfracdec[numberstops] = micro_stop_depth[i] = (float)(int)(micro_stop_depth[i]+0.499);
	  }
	}
       }
     }
    }
     k=0;
     if(numberstops>7) {
       k=numberstops/7;	   /* k=number of scrolls */
     }
     for(i=numberstops; i > 0; i--) {
       cont=0;
       if(i>numberstops) return 0; //i=numberstops;
       plotdepthdata(0, (numberstops-i), 5 );
       tissupdate((numberstops-i));
       if(ppo2 && (ppo2point[divenumber][numberstops-i+6]>0.16 || bailoutpoint[divenumber][numberstops-i+6]) ) {
	 ppo2fractionlast = ppo2point[divenumber][numberstops-i+6];
	 bailout = bailoutlast = bailoutpoint[divenumber][numberstops-i+6];
	 hemixlast = hemixpoint[divenumber][numberstops-i+6]+0.0001;
	 n2mixlast = n2mixpoint[divenumber][numberstops-i+6]+0.0001;
       }
       if(ppo2 && ppo2fractionlast<0.16) {
	 ppo2fractionlast=ppo2_limit_upper;
       }
       pambtolcalctrue();
       if(i<4 && !stoplookup_factor[sixstopmode][i]) {
	 nitrogenfracdec[i] = 0.00;
	 heliumfracdec[i] = 0.00;
	 ppo2fracdec[i] = 0.00;
	 hemixpointfracdec[i] = 0.0001;
	 n2mixpointfracdec[i] = 0.0001;
	 bailoutpointfracdec[i] = bailout;
	 stoptime[i] = 0.00;
	 continue;
       }
       if(!k) m=i;
       else {
	 m=i%7;
	 if(!m) {
	   m=7;
	   for(n=0; n<8; n++) {
	     _settextposition( n+4,46);
	     sprintf(stitle, "                                 ");
	     _outtext( stitle) ;
	   }
	 }
       }
       do {
	 oxtoolow=0;
	 currentstoppressure =
	   atmospheric + ( (((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )) *stopfactor) + decdepthinc) / 10.00 ;
	   //atmospheric + (((double)( ((i*stopfactor)==stopfactor && sixstopmode ) ? stopfactor*2.00 : (i*stopfactor) )) + decdepthinc) / 10.00 ;
	 absolutedepth = currentstoppressure;
	 absolutedepthpure =
	   atmospheric + ( (((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )) *stopfactor) + 0.00) / 10.00 ;
	   //((double)( ((i*stopfactor)==stopfactor && sixstopmode ) ? stopfactor*2.00 : (i*stopfactor) ))/10.00 + atmospheric;
	 if(microfracdec[i]) {
	   currentstoppressure=absolutedepthpure=absolutedepth=microfracdec[i]/10.0+atmospheric;
	 }
	 _settextposition( m+3,46);
	 sprintf(stitle, "                                ");
	 _outtext( stitle) ;
	 _settextposition( m+3,46);
	      sprintf(stitle, "For %3d%s%cstop",
		  (int)((absolutedepthpure-atmospheric)*10.00*feetfactor+0.49),
		  form,
		  ( ((i*stopfactor)!=(stoplookup_factor[sixstopmode][i]*stopfactor) && i<4 ) ? '*' : microfracdec[i]?'^' : ' '	)
		 ); //(int)( (((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )) *stopfactor) * feetfactor+0.49),
	 if(i<4 && stoplookup_factor[sixstopmode][i]==1.5 && feetfactor==1.00) { stitle[4]='4'; stitle[5]='.'; stitle[6]='5'; }
	 _outtext( stitle) ;
	 if(getgas( m+3,65, (absolutedepthpure-atmospheric)*10.00 )<0) {
	   _settextposition( m+3,46);
	    sprintf(stitle, "                                 ");
	   _outtext( stitle) ;
	   i+=2;
	   cont=1;
	   break;
	 }
	 nitrogenfracdec[i] = nitrogenfraction;
	 heliumfracdec[i] = heliumfraction;
	 ppo2fracdec[i] = ppo2fraction;
	 hemixpointfracdec[i] = hemix;
	 n2mixpointfracdec[i] = n2mix;
	 bailoutpointfracdec[i] = bailout;
	 stoptime[i] = 1.00;
	 tissuetotissuetemptransfer();
	 if(!microfracdec[i]) for(j=0; j<16; j++) {
	   stoptimetissue[j]=0.00;
	   tolstoppressure=atmospheric + (((double)((i<4 ? stoplookup_factor[sixstopmode][i-1] : i-1 )*stopfactor)) / 10.00);
	   pigttolstopminus1[j] = ( tolstoppressure ) / bcalc(j,toln2only,0) + acalc(j,toln2only,0);
	   if(!strcmp(argvg,"bignose2")) {
	     printf("\nptol-1=%g",pigttolstopminus1[j]);
	   }
	   fractioncalcs(i);
	   algorithmcalc(j, stoptime[i], i);
	   if( ( ((absolutedepth * heliumfractioncalc) < tissuetemphe[j]) || ((absolutedepth * nitrogenfractioncalc) < tissuetemp[j]) ) && (pigttolstopminus1[j] < (tissue[j] + tissuehe[j] ))) {
	     /*stoptimetissue[j] = ( halftime[j] / -0.69315) * log(1.0 - ((pigttolstopminus1[j] - tissue[j])/( releaserate * ((nitrogenfraction * absolutedepth) - tissue[j]) )));
	     */
	     stoptimetissue[j] = stoptimetisscalc( j, 1, stoptime[i]);
	     if(stoptimetissue[j] > stoptime[i]) {
	       stoptime[i] = stoptimetissue[j];
	     }
	     if(stoptime[i] == 29999.00) oxtoolow=1;
	   }
	   if(!strcmp(argvg,"bignose5")) {
	     printf("\nstiss=%g, stissh=%g",tissuetemp[j],tissuetemphe[j]);
	   }
	   if(!strcmp(argvg,"bignose4")) {
	     printf("\nAbsdepth=%g, exptime=%g  ",absolutedepth,exposuretime);
	   }
	 }
	 if(!strcmp(argvg,"bignose3")) {
	   printf("\nnf=%g, nh=%g, nfc=%g, nhc=%g",nitrogenfraction,heliumfraction,nitrogenfractioncalc,heliumfractioncalc);
	 }
	 if(oxtoolow) {
	   _settextposition( m+3,46);
	   sprintf(stitle, "WARNING: OXYGEN TOO LOW         ");
	   _outtext( stitle) ;
	   delay1sec();
	   if(air) {
	     _settextposition( m+3,46);
	     sprintf(stitle, "Decompression time too long     ");
	     _outtext( stitle) ;
	     delay1sec();
	     oxtoolow=0;
	   }
	 }
       } while (oxtoolow);
       if(!cont) {
	 if(microfracdec[i])
	   exposuretime = stoptime[i] = 1.00;
	 else {
	   absolutedepthpure = ((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )*stopfactor)/10.00 + atmospheric;
	   depth = (((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )*stopfactor) + decdepthinc);
	   exposuretime = stoptime[i];
	 }
	 /*
	 if(i!=numberstops) {
	   if(stoptime[i] < stoptime[i+1]) stoptime[i] = stoptime[i+1];
	 }
	 */
	 totaltimetosurface[divenumber] = totaltimetosurface[divenumber] + (long)(exposuretime + TIMEMOD); /*round up minutes*/
	 if(!strcmp(argvg,"bignose4")) {
	   printf("\nexptime=%g, depth=%g, absd=%g",exposuretime,depth,absolutedepth);
	   printf("\nnf=%g, nh=%g, nfc=%g, nhc=%g",nitrogenfraction,heliumfraction,nitrogenfractioncalc,heliumfractioncalc);
	   while(!kbhit()); getch();
	 }
	 //tissueupdate(0);
	 //currentcnsotudisplay(11, 3);
	 /* if(i!=1) */
	 _settextposition(m+3,62);
	 sprintf(stitle, "%3g + ___mins     ",(double) ( (int)(stoptime[i]+TIMEMOD) ));
	 _outtext( stitle) ;
	 if(!automaticmode && !autofinish) {
	   helpscreen(61);
	   sprintf(stitle, "%g", stoptimeplus[divenumber][i]);
	   _settextposition( m+3,68);
	   cnumbuf[0]=4;
	   numbuf = cgetsn( cnumbuf, "", stitle );
	   if(numbuf[0]>0x2f) stoptimeplus[divenumber][i] = (double)atoi( numbuf );
	   if(numbuf[0]==2) {
	     _settextposition( m+3,46);
	     sprintf(stitle, "                                 ");
	     _outtext( stitle) ;
	     i+=2;
	     continue;
	   }
	 }
	 _settextposition(m+3,62);
	 sprintf(stitle, "%3g + %gmins     ",(double) ( (int)(stoptime[i]+TIMEMOD) ),(double) ( (int)(stoptimeplus[divenumber][i]+TIMEMOD) ));
	 _outtext(stitle);
	 if(!strcmp(argvg,"bignosea")) printf("AAAAAAAAAAAAAAAAA");
	 plotdepthdata(0, (numberstops-i+1), 5 );
//hemixpoint[divenumber][6]=0.80;
//hemixpoint[divenumber][7]=0.80;
//hemixpoint[divenumber][8]=0.80;
//hemixpoint[divenumber][9]=0.80;
	 if(!strcmp(argvg,"bignosea")) printf("BBBBBBBBBBBBBBBBB");
	if(!strcmp(argvg,"bignosec")) {
	  printf("he%g", hemixpoint[divenumber][6] );
	}
	 tissupdate((numberstops-i+1));
	 plotdepthdata(0, (numberstops-i+1), 5 );
	 currentcnsotudisplay(11, 3);
	 if(!strcmp(argvg,"bignose4")) {
	   printf("\nexptime=%g, depth=%g, absd=%g",exposuretime,depth,absolutedepth);
	   printf("\nnf=%g, nh=%g, nfc=%g, nhc=%g",nitrogenfraction,heliumfraction,nitrogenfractioncalc,heliumfractioncalc);
	   while(!kbhit()); getch();
	 }
	 if(!strcmp(argvg,"bignose3")) {
	   printf("\nnf=%g, nh=%g, nfc=%g, nhc=%g",nitrogenfraction,heliumfraction,nitrogenfractioncalc,heliumfractioncalc);
	 }
       }
     }

     for(n=0; n<8; n++) {
       _settextposition( n+4,46);
       sprintf(stitle, "                                 ");
       _outtext( stitle) ;
     }

     k=0;
     if(numberstops>7) {
       k=numberstops/7;	   /* k=number of scrolls */
     }
     for(i=numberstops; i > 0; i--) {
       if(!k) m=i;
       else {
	 m=i%7;
	 if(!m){
	   m=7;
	   if( (i>6) && i!=numberstops) {
	     _settextposition( 11,46);
	     sprintf(stitle, "Press any key for next page.... ");
	     _outtext( stitle) ;
	     helpscreen(10);
	     getch();
	   }
	   for(n=0; n<8; n++) {
	     _settextposition( n+4,46);
	     sprintf(stitle, "                                 ");
	     _outtext( stitle) ;
	   }
	 }
       }
       _settextposition( m+3,46);
       sprintf(stitle, "                                ");
       _outtext( stitle) ;
       _settextposition( m+3,46);
       if(i<4 && !stoplookup_factor[sixstopmode][i]) {
	 continue;
       }
	   //(microfracdec[i] ? (int)microfracdec[i] : (int)(((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )*stopfactor))) * feetfactor+0.49,
       if(microfracdec[i]) {
	 sprintf(stitle, "%3d%s%cstop: %2.0fmins: ",
	   (int)(microfracdec[i]*feetfactor+0.49),
	   form,
	   '^',
	   (double) ( (int)(stoptime[i]+TIMEMOD+stoptimeplus[divenumber][i]) ) );
       }
       else {
	 sprintf(stitle, "%3d%s%cstop: %2.0fmins: ",
	   (int)(((double)(i<4 ? stoplookup_factor[sixstopmode][i] : i )*stopfactor)*feetfactor+0.49),
	   form,
	   ( ((i*stopfactor)!=(stoplookup_factor[sixstopmode][i]*stopfactor) && i<4 ) ? '*' : ' ' ),
	   (double) ( (int)(stoptime[i]+TIMEMOD+stoptimeplus[divenumber][i]) ) );
	 if(i<4 && stoplookup_factor[sixstopmode][i]==1.5 && feetfactor==1.00) { stitle[0]='4'; stitle[1]='.'; stitle[2]='5'; }
       }
       _outtext( stitle) ;
       if(ppo2 && !bailoutpointfracdec[i]) {
	 ppo2print(ppo2fracdec[i]);
	 sprintf(stitle, "PPO2=%1d.%02dbar", ppi, ppf);
	 _outtext( stitle) ;
       }
       else {
	 if(air || n2) {
	   if(heo2display) sprintf(stitle, "%d%%O2 ", (int)(OXYGENFRACDECI*100.00+0.49) );
	   else sprintf(stitle, "%d%%N2 ", (int)(nitrogenfracdec[i]*100.00+0.49) );
	   _outtext( stitle) ;
	 }
	 if(he) {
	   sprintf(stitle, "%d%%He", (int)(heliumfracdec[i]*100.00+0.49) );
	   _outtext( stitle) ;
	 }
       }
     }
  return 1;
}


void plotdepthdata(int fullrun, int stopsprocessed, int depthprocessed)
{
int numtime, numdepth, snum, i, j, x, y, p, c, oxtoolow, m;
double depthmax=0.00, timetotal=0.00, ddiff, dlast=0, tlast, timepointtotal=0;
unsigned char title[100];

  if( vc.numcolors > 2)_setcolor(0);
  _setviewport( x1graph1/pixfact,y1graph1/piyfact, x2graph1/pixfact,y2graph1/piyfact );
  borderdraw();
  _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact);
  if( vc.numcolors > 2)_setcolor(8);
  x=(x1graph1+6)/pixfact; y=y1graph1/piyfact;
  _rectangle( _GFILLINTERIOR, x+(x2graph1-x1graph1-6)/pixfact, y+(y2graph1-y1graph1-6)/piyfact, x, y );

  for(i=0; i<=depthprocessed; i++) {
    depthpoint[divenumber][i] = depthdepth[i];
    timepointb[divenumber][i] = exposuretimedepth[i];
    timepointc[divenumber][i] = 0.00;
    nitrogenpoint[divenumber][i] = nitrogenfractiondepth[i];
    heliumpoint[divenumber][i] = heliumfractiondepth[i];
    ppo2point[divenumber][i] = ppo2fractiondepth[i];
    hemixpoint[divenumber][i] = hemixpointfractiondepth[i]+0.0001;
    n2mixpoint[divenumber][i] = n2mixpointfractiondepth[i]+0.0001;
    bailoutpoint[divenumber][i] = bailoutpointfractiondepth[i];
    ddiff = fabs(dlast - depthpoint[divenumber][i]);
    if( (dlast - depthpoint[divenumber][i]) < 0 ) timepointa[divenumber][i] = ddiff / DESCENTRATE;
    else timepointa[divenumber][i] = ascenttimediff(dlast, depthpoint[divenumber][i]);
    dlast = depthpoint[divenumber][i];
  }
  i=6;
  for(j=numberstops ; j>(numberstops-stopsprocessed) ; i++, j--) {
    if(microfracdec[j])
      depthpoint[divenumber][i] = microfracdec[j];
    else depthpoint[divenumber][i] =
      (double)( j<4 ? stopfactor*stoplookup_factor[sixstopmode][j] : j*stopfactor );
      //(double)( ((j*stopfactor)==stopfactor  && sixstopmode ) ? stopfactor*2.00 : (j*stopfactor) );
    timepointb[divenumber][i] = (double)( (int)(stoptime[j]+TIMEMOD) );
    timepointc[divenumber][i] = stoptimeplus[divenumber][j];
    nitrogenpoint[divenumber][i] = nitrogenfracdec[j];
    heliumpoint[divenumber][i] = heliumfracdec[j];
    ppo2point[divenumber][i] = ppo2fracdec[j];
    hemixpoint[divenumber][i] = hemixpointfracdec[j]+0.0001;
    n2mixpoint[divenumber][i] = n2mixpointfracdec[j]+0.0001;
    bailoutpoint[divenumber][i] = bailoutpointfracdec[j];

    ddiff = fabs(dlast - depthpoint[divenumber][i]);
    if( (dlast - depthpoint[divenumber][i]) < 0 ) timepointa[divenumber][i] = ddiff / DESCENTRATE;
    else timepointa[divenumber][i] = ascenttimediff(dlast, depthpoint[divenumber][i]);
    dlast = depthpoint[divenumber][i];
  }
  /* Clear out next point zero time, with last point gas data */
  depthpoint[divenumber][i] = 0.00;
  timepointa[divenumber][i] = 0.00;
  timepointb[divenumber][i] = 0.00;
  nitrogenpoint[divenumber][i] = nitrogenpoint[divenumber][i-1];
  heliumpoint[divenumber][i] = heliumpoint[divenumber][i-1];
  //ppo2point[divenumber][i] = ppo2point[divenumber][i-1];
  //hemixpoint[divenumber][i] = hemixpoint[divenumber][i-1];
  //n2mixpoint[divenumber][i] = n2mixpoint[divenumber][i-1];
  //bailoutpoint[divenumber][i] = bailoutpoint[divenumber][i-1];
  ddiff = fabs(dlast - depthpoint[divenumber][i]);
  if( (dlast - depthpoint[divenumber][i]) < 0 ) timepointa[divenumber][i] = ddiff / DESCENTRATE;
  else timepointa[divenumber][i] = ascenttimediff(dlast, depthpoint[divenumber][i]);
  dlast = depthpoint[divenumber][i];

  numberpoints[divenumber] = i;

  for(i=0 ;i<=numberpoints[divenumber]; i++) {
    if(i==depthprocessed+1) i=6;
    if(depthmax < depthpoint[divenumber][i]) depthmax = depthpoint[divenumber][i];
    timetotal = timetotal + timepointb[divenumber][i]+ timepointa[divenumber][i] + timepointc[divenumber][i] ;
  }

  titleprofilegraph(xposgraph1, depthmax, timetotal );

  //  if( vc.numcolors > 2)_setcolor(14);
  _setviewport( (x1graph1+5)/pixfact,y1graph1/piyfact, x2graph1/pixfact,(y2graph1-5)/piyfact );

  _moveto_w( 0.00/(double)pixfact, 0.00/(double)piyfact);
  for(i=0 ;i<=(numberpoints[divenumber] - (fullrun ? 0 : 1)) ; i++) {
      if(i==depthprocessed+1) i=6;
      if(bailoutpoint[divenumber][i]) {
	if( vc.numcolors > 2) _setcolor(14);
      }
      else {
	if( vc.numcolors > 2) _setcolor(10);
      }
      //printf("bailout%d",bailoutpoint[divenumber][i]);
      //getch();
    _lineto_w( (timepointtotal+timepointa[divenumber][i])*timetotalgraph, depthpoint[divenumber][i]*depthmaxgraph);
    timepointtotal =  timepointtotal + timepointa[divenumber][i];
    _lineto_w( (timepointtotal+timepointb[divenumber][i]+timepointc[divenumber][i] )*timetotalgraph, depthpoint[divenumber][i]*depthmaxgraph);
    timepointtotal =  timepointtotal + timepointb[divenumber][i] + timepointc[divenumber][i] ;
  }

  tissuegraph();
       if(!strcmp(argvg,"bignose5")) {
	 while(!kbhit()); getch();
       }
  if(fullrun) {
    ppo2cns[divenumber] = ppo2cnscurrent;
    for(m=8;m<11;m++) {
      _settextposition( m,3);
      sprintf(stitle, "                                          ");
      _outtext( stitle) ;
    }
    _settextposition( 7,3);
    sprintf(stitle, "Time to surface=%ldmins(stop1 in %ldmins)", totaltimetosurface[divenumber], timetofirststop[divenumber]);
    _outtext( stitle) ;
    _settextposition( 8,3);
    if( (ppo2cnsmax[divenumber]>100.00) && (vc.numcolors > 2) )_settextcolor(4);
    sprintf(title, "CNS exposure: %d%%peak, %d%%dive end ", (int)ppo2cnsmax[divenumber], (int)ppo2cns[divenumber]);
    _outtext(title);
    if( vc.numcolors > 2) _settextcolor(7);
    _settextposition( 9,3);
    ppo2print(maxppo2[divenumber]);
    if( (maxppo2[divenumber]>2.00) && (vc.numcolors > 2) )_settextcolor(4);
    sprintf(stitle, "MaxPPO2=%1d.%02dbar", ppi, ppf);
    _outtext( stitle) ;
    if( vc.numcolors > 2) _settextcolor(7);
    sprintf(stitle, "  OTU=%d  OTUtotal=%d", (int)diveotu[divenumber], (int)missionotu);
    _outtext( stitle) ;

    divefinish[divenumber][0] = missiontotal[0];
    divefinish[divenumber][1] = missiontotal[1];
    divefinish[divenumber][2] = missiontotal[2];
    _setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact );
    settitletext();
    if( vc.numcolors > 2)_setcolor(0);
    _moveto( 20/pixfact, 26/piyfact);
    sprintf(title,"Dive Finish:  Day%2d   Time %02d:%02d  ", divefinish[divenumber][0], divefinish[divenumber][1], divefinish[divenumber][2]);
    _outgtext( title);

    do {
      oxtoolow=0;
      nitrogenfractionlast=0.79;
      heliumfractionlast=0.001;
      oxygenfractionlast=0.21;
      _settextposition( 11,3);
      sprintf(stitle, "                                          ");
      _outtext( stitle) ;
      _settextposition( 10,3);
      sprintf(stitle, "Surface gas               ");
      _outtext( stitle) ;
      while(getgas(10,15,0.00)<0);
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
	  stoptimetissue[j] = stoptimetisscalc( j, 0, flytolmins);
	  if(stoptimetissue[j] > flytolmins)  {
	    flytolmins = stoptimetissue[j];
	  }
	  //if(stoptimetissue[j] == 29999.00) oxtoolow=1;
	}
      }
      if(oxtoolow) {
	_settextposition( 10,3);
	sprintf(stitle, "WARNING: OXYGEN TOO LOW   ");
	_outtext( stitle) ;
	delay1sec();
      }
    } while (oxtoolow);
    flytolupdate();
    if( vc.numcolors > 2)_setcolor(0);
    settitletext();
    _moveto( 320/pixfact, 11/piyfact);
    sprintf(title,"Flight Time:  Day%2d   Time %02d:%02d  ", flytol[0], flytol[1], flytol[2]);
    _outgtext( title);
    helpscreen(16);
    _settextposition( 11,3);
    sprintf(stitle, "Print screen? Y/<N>                       ");
    _outtext( stitle) ;
    c=getchyn();
    if(c=='y' || c=='Y') {
      _settextposition( 11,3);
      sprintf(stitle, "                        ");
      _outtext( stitle) ;
      _setviewport( 0, 0, vc.numxpixels, vc.numypixels);
      graphicscreenprint();
    }
    helpscreen(7);
    _settextposition( 11,3);
    sprintf(stitle, "Surface time ____mins                     ");
    _outtext( stitle) ;
    _settextposition( 11,16);
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
    timepointc[divenumber][i]=0.00;
    depthpoint[divenumber][i] = depth;
    timepointb[divenumber][i] = surftime;
    timepointc[divenumber][i]=0.00;
    nitrogenpoint[divenumber][i] = nitrogenfraction;
    heliumpoint[divenumber][i] = heliumfraction;
    ppo2point[divenumber][i] = ppo2fraction;
    hemixpoint[divenumber][i] = hemix+0.0001;
    n2mixpoint[divenumber][i] = n2mix+0.0001;
    bailoutpoint[divenumber][i] = bailout;
    ddiff = fabs(dlast - depthpoint[divenumber][i]);
    if( (dlast - depthpoint[divenumber][i]) < 0 ) timepointa[divenumber][i] = ddiff / DESCENTRATE;
    else timepointa[divenumber][i] = ascenttimediff(dlast, depthpoint[divenumber][i]);
    dlast = depthpoint[divenumber][i];
    tissueupdate(0);
  }

}


int dispaydepthdata(void)
{
int i, j, k, m, ii, step, tablemix;
unsigned char title[100], titlednum[10];
int jj;
  repeat=1;
  for (i=0,j=3,m=4; i<6; step=0) {
   m=4+ 4*(i/3);
   j=3+(i%3)*14;
   do {
    if(i) helpscreen(2);
    else helpscreen(1);
    do {
	_settextposition( m,j);
	sprintf(stitle, "Depth%d=___%s", i+1, form);
	_outtext( stitle) ;
	if(repeat) {
	  depthdepth[i] = 0.00;
	  if(timepointb[divenumber][i]) {
	    depthdepth[i] = depthpoint[divenumber][i];
	    nitrogenfractionlast = nitrogenpoint[divenumber][i];
	    heliumfractionlast = heliumpoint[divenumber][i];
	    if(ppo2 && (ppo2point[divenumber][i]>0.16 || bailoutpoint[divenumber][i]) ) {
	      ppo2fractionlast = ppo2point[divenumber][i];
	      bailout = bailoutlast = bailoutpoint[divenumber][i];
	      hemixlast = hemixpoint[divenumber][i]+0.0001;
	      n2mixlast = n2mixpoint[divenumber][i]+0.0001;
	    }
	  }
	  sprintf(stitle, "%d", (int)(depthdepth[i] * feetfactor+0.49));
	}
	_settextposition( m,j+7);
	cnumbuf[0]=4;
	numbuf = cgetsn( cnumbuf, "", stitle );
	if(numbuf[0]==2) {
	  if(i==0) { break; }//numbuf[0]='0'; numbuf[1]=0; }
	  else {
	    step=-1;
	    break;
	  }
	}
      //for (ii=i ;ii<6;ii++) depthdepth[ii] = (double)atoi( numbuf ) / feetfactor;
      depthdepth[i] = (double)atoi( numbuf ) / feetfactor;
      if(repeat && numbuf[0]<0x10) {
	if(!depthdepth[i] && timepointb[divenumber][i]) depthdepth[i] = depthpoint[divenumber][i];
      }
      depthpoint[divenumber][i] = depthdepth[i];
      if(!depthdepth[i]) {
	if( i) {
	  _settextposition( m,j);
	  sprintf(stitle, "             ");
	  _outtext( stitle) ;
	  for(;i<6;i++) {
	    depthdepth[i] = depthdepth[i-1];
	    exposuretimedepth[i] = 0.00;
	    nitrogenfractiondepth[i] = nitrogenfractiondepth[i-1];
	    heliumfractiondepth[i] = heliumfractiondepth[i-1];
	    ppo2fractiondepth[i] = ppo2fractiondepth[i-1];
	    hemixpointfractiondepth[i] = hemixpointfractiondepth[i-1];
	    n2mixpointfractiondepth[i] = n2mixpointfractiondepth[i-1];
	    bailoutpointfractiondepth[i] = bailoutpointfractiondepth[i-1];

	  }
	return 0;
	}
	else return 1; /*abort*/
      }
      if(!i) {
	divestart[divenumber][0] = missiontotal[0];
	divestart[divenumber][1] = missiontotal[1];
	divestart[divenumber][2] = missiontotal[2];
	ppo2exptime_14[divenumber]=0.00;
	ppo2exptime_15[divenumber]=0.00;
	ppo2exptime_16[divenumber]=0.00;
	ppo2exptime_16plus[divenumber]=0.00;
	maxppo2[divenumber]=0.00;
	diveotu[divenumber]=0.00;
	ppo2cns[divenumber]=ppo2cnscurrent;
	ppo2cnsmax[divenumber]=ppo2cnscurrent;
	sixstopmodedive=sixstopmode;
	numberstops=0;
	for(k=0; k<50; k++) dailyotu[k]=0.00;

	_setviewport( 0/pixfact,0/piyfact, 640/pixfact,480/piyfact);
	settitletext();
	if( vc.numcolors > 2)_setcolor(0);
	_moveto( 20/pixfact, 11/piyfact);
	sprintf(title,"Dive Start:     Day%2d   Time %02d:%02d  ", divestart[divenumber][0], divestart[divenumber][1], divestart[divenumber][2]);
	_outgtext( title);
      }
      if(maxdepthalarm < depthdepth[i]) helpscreen(8);
      //absolutedepthpure = ((double)depthdepth[i]/10.00) + atmospheric;
      //depth = (depthdepth[i] * 1.03) + depthinc;
      //exposuretime = 0.001;
      //tissueupdate(1); // Recalc pambtolmindepth after ascent
      pambtolcalctrue();
      if(depthdepth[i] < pambtolmindepth) helpscreen(9);
      step=1;
    } while( (maxdepthalarm < depthdepth[i]) || (depthdepth[i] < pambtolmindepth) );
    if(step==1) {
      _settextposition( m,j);
      sprintf(stitle, "Depth%d=%d%s  ", i+1, (int)(depthdepth[i] * feetfactor+0.49), form);
      _outtext( stitle) ;
    }
    if(step==1) do {
      helpscreen(3);
      _settextposition( m+1,j);
      sprintf(stitle, "Time=___mins  ");
      _outtext( stitle) ;
      if(repeat) {
	exposuretimedepth[i]=timepointb[divenumber][i];
	sprintf(stitle, "%g", exposuretimedepth[i]);
      }
      _settextposition( m+1,j+5);
      cnumbuf[0]=4;
      numbuf = cgetsn( cnumbuf, "", stitle );
      if(numbuf[0]==2) { step=0; break; }
      exposuretimedepth[i] = (double)atoi( numbuf );
      if(repeat) {
	if(!exposuretimedepth[i]) exposuretimedepth[i]=timepointb[divenumber][i];
      }
      if(!exposuretimedepth[i]) {
	exposuretimedepth[i] = 1.00;
      }
      timepointb[divenumber][i] = exposuretimedepth[i];
      timepointc[divenumber][i] = 0.00;
      step=2;
    } while ( !exposuretimedepth[i] && (!i || ((depthdepth[i] - depthdepth[i-1]) > 0.00)) );
    if(step==2) {
      _settextposition( m+1,j);
      sprintf(stitle, "Time=%gmins  ", exposuretimedepth[i]);
      _outtext( stitle) ;
      if(step==2) {
	if(getgas(m+2,j,depthdepth[i]) <0) { step=1; break;}
      }
      absolutedepthpure = ((double)depthdepth[i]/10.00) + atmospheric;
      depth = (depthdepth[i] * 1.03) + depthinc;
      exposuretime = exposuretimedepth[i];
      for (ii=i ;ii<6;ii++) {
	nitrogenfractiondepth[ii] = nitrogenfraction;
	heliumfractiondepth[ii] = heliumfraction;
	ppo2fractiondepth[ii] = ppo2fraction;
	hemixpointfractiondepth[ii] = hemix;
	n2mixpointfractiondepth[ii] = n2mix;
	bailoutpointfractiondepth[ii] =	bailout;
      }
      if(i<5) {
	for (ii=i+1 ;ii<6;ii++) exposuretimedepth[ii] = 0.00;
      }
      step=3;
    }
    if(step==-1 && i) {
      _settextposition( m,j);
      sprintf(stitle, "              ");
      _outtext( stitle) ;
      _settextposition( m+1,j);
      sprintf(stitle, "              ");
      _outtext( stitle) ;
      _settextposition( m+2,j);
      sprintf(stitle, "              ");
      _outtext( stitle) ;
      i-=2;
    }
    if(step==-1 || step==3) {
      tissueorgtransfer();
      depthlast=0.00;
      for(ii=0; ii<=i && timepointb[divenumber][ii] && depthpoint[divenumber][ii]; ii++) {
	depth=(depthpoint[divenumber][ii] * 1.03) + depthinc;
	absolutedepthpure=depthpoint[divenumber][ii]/10.00 + atmospheric;
	exposuretime=timepointb[divenumber][ii];
	nitrogenfraction=nitrogenfractiondepth[ii]; //nitrogenpoint[divenumber][
	heliumfraction=heliumfractiondepth[ii];  //heliumpoint[divenumber][
	hemix = hemixpointfractiondepth[ii]+0.0001;
	n2mix = n2mixpointfractiondepth[ii]+0.0001;
	tissueupdate(0);
      }
      currentcnsotudisplay(11, 3);
      plotdepthdata(0, numberstops, i);
      i++;
      step=3;
    }
   }while(step<3);
	 if(i==1) {
	  absolutedepth=(absolutedepthpure - atmospheric)*1.03 +0.1 + atmospheric;
	  stoptime[i]=999.00;
	  for(jj=0; jj<16; jj++) {
	   stoptimetissue[jj]=999.00;
	   tolstoppressure=atmospheric; // + (((double)((i<4 ? stoplookup_factor[sixstopmode][i-1] : i-1 )*stopfactor)) / 10.00);
	   pigttolstopminus1[jj] = ( tolstoppressure ) / bcalc(jj,toln2only,0) + acalc(jj,toln2only,0);
	   fractioncalcs(30);
	   algorithmcalc(jj, 0.00, 1);
	   if(	(pigttolstopminus1[jj] > (tissue[jj] + tissuehe[jj] ))) {
	   //if( ( ((absolutedepth * heliumfractioncalc) < tissuetemphe[jj]) || ((absolutedepth * nitrogenfractioncalc) < tissuetemp[jj]) ) && (pigttolstopminus1[jj] < (tissue[jj] + tissuehe[jj] ))) {
	     /*stoptimetissue[jj] = ( halftime[jj] / -0.69315) * log(1.0 - ((pigttolstopminus1[jj] - tissue[jj])/( releaserate * ((nitrogenfraction * absolutedepth) - tissue[jj]) )));
	     */
	     stoptimetissue[jj] = stoptimetisscalc( jj, 1, 5.00);
	     if(stoptimetissue[jj] < stoptime[i]) {
	       stoptime[i] = stoptimetissue[jj];
	     }
	     //if(stoptime[i] == 29999.00) oxtoolow=1;
	   }
	   else break;
	  }
	  printf("\nNST=%f",stoptime[i]);
	  getch();
	 }

  }
  return 0;
}

void currentcnsotudisplay(int x, int y)
{
    if( vc.numcolors > 2) _settextcolor(14);
    _settextposition( x, y);
    ppo2print(ppo2_now);
    sprintf(stitle, "Current CNS=%d%%, OTU=%d, PPO2=%1d.%02dbar  ", (int)ppo2cnscurrent, (int)diveotu[divenumber], ppi, ppf);
    _outtext( stitle) ;
    if( vc.numcolors > 2) _settextcolor(7);
}


void tradecylinderfill(void)
{
int i, c, heinc=0, topup=0, entergaspercent=0;
double sftemp=0.00, fillpressure=0.00, workingdepth=0.00, workingppo2=0.00, o2frac=0.00, n2frac=0.00, hefrac=0.00, narcpressure=0.00, hefill=0.00, airfill=0.00, o2fill=0.00, n2fill=0.00;
char message[MAXSTR];
unsigned char title[100], row;
double hefillpressure, o2fillpressure, hefillfrac, o2fillfrac, n2fillfrac, airtopoff;

  drawbackground();
  _moveto( 75/pixfact, 11/piyfact);
  _outgtext("TRADE CYLINDER FILLING");
  _moveto( 75/pixfact, 26/piyfact);
  _outgtext("     CALCULATIONS");
  if( vc.numcolors > 2) _setcolor(7);
  row=8;

  do {
    helpscreen(44);
    _settextposition( row,20);
    sprintf(stitle, "Maximum working cylinder pressure ____%s",porb);
    _outtext( stitle) ;
    _settextposition( row++,54);
    cnumbuf[0]=5;
    numbuf = cgetsn( cnumbuf, "", "" );
    if(!*numbuf) return;
    sftemp = ( (double)atof( numbuf ) ) ;
  } while( (sftemp < 0.00) || !*numbuf);
  trademaxcylinderpressure = sftemp/psifactor;
  if(feetfactor==1.00) {
    do {
      helpscreen(45);
      _settextposition( row,20);
      sprintf(stitle, "Cylinder size = ____litres                  ");
      _outtext( stitle) ;
      _settextposition( row++,36);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atof( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    tradecylindersize = sftemp;
  }
  else {
    do {
      helpscreen(46);
      _settextposition( row,20);
      sprintf(stitle, "Free air capacity of cylinder = ____cubic feet");
      _outtext( stitle) ;
      _settextposition( row++,52);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atof( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    tradefreecylindersize = sftemp*cuft_ltr_factor;
    tradecylindersize = tradefreecylindersize / trademaxcylinderpressure;
  }

  helpscreen(29);
  _settextposition( row++,20);
  sprintf(stitle, "Is air top off only to cylinder required? Y/<N> ");
  _outtext( stitle) ;
  do {
    topup=getch();
    if(topup==27) return;
    if(topup==13 || topup==10) topup='N';
    if(topup!='y' && topup!='Y' && topup!='n' && topup!='N') printf("%c",7);
  }while(topup!='y' && topup!='Y' && topup!='n' && topup!='N');
  if( vc.numcolors > 2)_settextcolor(15);
  if(topup=='y' || topup=='Y') {
    sprintf(stitle, "Y");
    _outtext( stitle) ;
    topup=1;
  }
  else {
    sprintf(stitle, "N");
    _outtext( stitle) ;
    topup=0;
  }
  helpscreen(30);
  if( vc.numcolors > 2)_settextcolor(7);
  _settextposition( row++,20);
  if(!topup) sprintf(stitle, "Is helium used in mix Y/<N> ");
  else sprintf(stitle, "Is helium used in current mix Y/<N> ");
  _outtext( stitle) ;
  do {
    heinc=getch();
    if(heinc==27) return;
    if(heinc==13 || heinc==10) heinc='N';
    if(heinc!='y' && heinc!='Y' && heinc!='n' && heinc!='N') printf("%c",7);
  }while(heinc!='y' && heinc!='Y' && heinc!='n' && heinc!='N');
  if( vc.numcolors > 2)_settextcolor(15);
  if(heinc=='y' || heinc=='Y') {
    sprintf(stitle, "Y");
    _outtext( stitle) ;
    heinc=1;
  }
  else {
    sprintf(stitle, "N");
    _outtext( stitle) ;
    heinc=0;
  }
  if( vc.numcolors > 2)_settextcolor(7);
  if(topup) {
    helpscreen(31);
    _settextposition( row++,20);
    sprintf(stitle, "Do you wish to enter current gas as ");
    _outtext( stitle) ;
    if( vc.numcolors > 2)_settextcolor(15);
    sprintf(stitle, "P");
    _outtext( stitle) ;
    if( vc.numcolors > 2)_settextcolor(7);
    sprintf(stitle, "ercentages");
    _outtext( stitle) ;
    _settextposition( row++,20);
    sprintf(stitle, "or as p");
    _outtext( stitle) ;
    if( vc.numcolors > 2)_settextcolor(15);
    sprintf(stitle, "R");
    _outtext( stitle) ;
    if( vc.numcolors > 2)_settextcolor(7);
    sprintf(stitle, "ressures?  P/R... ");
    _outtext( stitle) ;
    do {
      entergaspercent=getch();
      if(entergaspercent==27) return;
      if(entergaspercent!='p' && entergaspercent!='P' && entergaspercent!='r' && entergaspercent!='R') printf("%c",7);
    }while(entergaspercent!='p' && entergaspercent!='P' && entergaspercent!='r' && entergaspercent!='R');
    if( vc.numcolors > 2)_settextcolor(15);
    if(entergaspercent=='p' || entergaspercent=='P') {
      sprintf(stitle, "P");
      _outtext( stitle) ;
      entergaspercent=1;
    }
    else {
      sprintf(stitle, "R");
      _outtext( stitle) ;
      entergaspercent=0;
    }
    if( vc.numcolors > 2)_settextcolor(7);
  }

  if(topup && !entergaspercent) {
    helpscreen(32);
    do {
      do {
	/* helpscreen(33); */
       _settextposition(row,20);
       sprintf(stitle, "Enter current O2 Fill pressure= _____%s",porb);
       _outtext( stitle) ;
       _settextposition(row++,52);
       cnumbuf[0]=6;
       numbuf = cgetsn( cnumbuf, "", "" );
       if(!*numbuf) return;
       sftemp = ( (double)atof( numbuf ) ) ;
      } while( (sftemp < 0.00) || !*numbuf);
      o2fillpressure = sftemp/psifactor;
      if(heinc) {
	do {
	  _settextposition(row,20);
	  sprintf(stitle, "Enter current He Fill pressure= _____%s",porb);
	  _outtext( stitle) ;
	  _settextposition(row++,52);
	  cnumbuf[0]=6;
	  numbuf = cgetsn( cnumbuf, "", "" );
	  if(!*numbuf) return;
	  sftemp = ( (double)atof( numbuf ) ) ;
	} while( (sftemp < 0.00) || !*numbuf);
	hefillpressure = sftemp/psifactor;
      }
      else hefillpressure = 0.00;
      currentfillpressure = hefillpressure + o2fillpressure;
      if(currentfillpressure>trademaxcylinderpressure) {
	printf("%c",7);
	helpscreen(47);
	row--;
	row--;
      }
    }while(currentfillpressure>trademaxcylinderpressure);
  }
  if(topup && entergaspercent && heinc) {
    do {
      helpscreen(34);
      _settextposition(row,20);
      sprintf(stitle, "Enter current He Fill percent= ____%%");
      _outtext( stitle) ;
      _settextposition(row++,51);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atoi( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    hefillfrac = sftemp/100.00;
  }
  else hefillfrac = 0.00;
  if(topup && entergaspercent) {
    do {
      helpscreen(35);
      _settextposition(row,20);
      sprintf(stitle, "Enter current O2 Fill percent= ____%%");
      _outtext( stitle) ;
      _settextposition(row++,51);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atoi( numbuf ) ) ;
    } while( (sftemp < 0.00) || !*numbuf);
    o2fillfrac = sftemp/100.00;
    n2fillfrac = 1.00 - o2fillfrac - hefillfrac;
    do {
      helpscreen(36);
      _settextposition(row,20);
      sprintf(stitle, "Enter current cylinder pressure= ____%s",porb);
      _outtext( stitle) ;
      _settextposition(row++,53);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atof( numbuf ) ) ;
      if(sftemp/psifactor>trademaxcylinderpressure) {
	printf("%c",7);
	helpscreen(47);
	row--;
      }
    } while( (sftemp < 0.00) || !*numbuf || (sftemp/psifactor>trademaxcylinderpressure) );
    currentfillpressure = sftemp/psifactor;
  }

  if(!topup) {
    helpscreen(27);
    do {
      _settextposition( row, 20);
      sprintf(stitle, "Enter O2=__%%");
      _outtext( stitle) ;
      _settextposition( row++, 29);
      o2frac = fillgetgas();

      if(heinc) {
	_settextposition( row, 20);
	sprintf(stitle, "Enter He=__%%");
	_outtext( stitle) ;
	_settextposition( row++, 29);
	hefrac = fillgetgas();
      }
      else hefrac=0.00;

      n2frac = 1.00 - hefrac - o2frac;
      if( n2frac<0.00 || n2frac>0.79001 || (o2frac/(o2frac+n2frac))<0.20999 ) {
	if(heinc) {
	  row--;
	  _settextposition( row, 20);
	  sprintf(stitle, "            ");
	  _outtext( stitle) ;
	}
	row--;
	_settextposition( row, 20);
	sprintf(stitle, "            ");
	_outtext( stitle) ;
	if(n2frac<0.00)helpscreen(28);
	if( n2frac>0.79001 || (o2frac/(o2frac+n2frac))<0.20999 ) helpscreen(43);
      }
    } while(n2frac<0.00 || n2frac>0.79001 || (o2frac/(o2frac+n2frac))<0.20999 );
  }

  if(topup) {
    helpscreen(37);
    do {
      _settextposition(row,20);
      sprintf(stitle, "Enter new required fill pressure= ____%s",porb);
      _outtext( stitle) ;
      _settextposition(row++,54);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atof( numbuf ) ) ;
      fillpressure = sftemp/psifactor;
      airtopoff=0.00;
      if(entergaspercent) {
	airtopoff = fillpressure - currentfillpressure;
	hefrac = hefillfrac*currentfillpressure / fillpressure;
	o2frac = ( o2fillfrac*currentfillpressure + 0.21*airtopoff ) / fillpressure;
	n2frac = 1.00 - hefrac - o2frac;
      }
      else {
	airtopoff = fillpressure - o2fillpressure - hefillpressure;
	hefrac = hefillpressure / fillpressure;
	o2frac = ( o2fillpressure + 0.21*airtopoff ) / fillpressure;
	n2frac = 1.00 - hefrac - o2frac;
      }
      if(airtopoff<0.00) {
	printf("%c",7);
	helpscreen(40);
	row--;
      }
      if(fillpressure>trademaxcylinderpressure) {
	printf("%c",7);
	helpscreen(47);
	row--;
      }
    } while( (sftemp < 0.00) || !*numbuf || (airtopoff<0.00) || (fillpressure>trademaxcylinderpressure) );
  }
  else {
    helpscreen(38);
    do {
      _settextposition(row,20);
      sprintf(stitle, "Enter required fill pressure= ____%s",porb);
      _outtext( stitle) ;
      _settextposition(row++,50);
      cnumbuf[0]=5;
      numbuf = cgetsn( cnumbuf, "", "" );
      if(!*numbuf) return;
      sftemp = ( (double)atof( numbuf ) ) ;
      if(sftemp/psifactor>trademaxcylinderpressure) {
	printf("%c",7);
	helpscreen(47);
	row--;
      }
    } while( (sftemp < 0.00) || !*numbuf || (sftemp/psifactor>trademaxcylinderpressure) );
    fillpressure = sftemp/psifactor;
  }
  row++;
  if( vc.numcolors > 2)_settextcolor(15);
  _settextposition( row++,20);
  sprintf(stitle, "Final gas mix: ");
  _outtext( stitle) ;
  if(o2frac>0.00) {
    sprintf(stitle, "%.1f%%O2, ", (o2frac*100.00) );
    _outtext( stitle) ;
  }
  if(hefrac>0.00) {
    sprintf(stitle, "%.1f%%He, ", (hefrac*100.00) );
    _outtext( stitle) ;
  }
  if(n2frac>0.00) {
    sprintf(stitle, "%.1f%%N2", (n2frac*100.00) );
    _outtext( stitle) ;
  }
  if(!topup) {
    _settextposition( row++,20);
    if(heinc) {
      hefill = hefrac * fillpressure;
      airfill = ( n2frac * fillpressure ) / 0.79 ;
      o2fill = fillpressure - hefill - airfill;
      hefillltr=hefill*tradecylindersize;
      hefillprice=hefillltr*heprice;
      o2fillltr=o2fill*tradecylindersize;
      o2fillprice=o2fillltr*o2price;
      airfillltr=airfill*tradecylindersize;
      airfillprice=airfillltr*airprice;
      sprintf(stitle, "Oxygen fill: %.1f%s, %.1f%s, %s%.3f", o2fill*psifactor, porb, o2fillltr/cuft_ltr_factor, cuftorltr, gascurrency, o2fillprice);
      _outtext( stitle) ;
      _settextposition( row++,20);
      sprintf(stitle, "Helium fill: %.1f%s, %.1f%s, %s%.3f", hefill*psifactor, porb, hefillltr/cuft_ltr_factor, cuftorltr, gascurrency, hefillprice);
      _outtext( stitle) ;
      _settextposition( row++,20);
      sprintf(stitle, "Air fill   : %.1f%s, %.1f%s, %s%.3f", airfill*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
      _outtext( stitle) ;
    }
    else {
      if(o2frac<0.21) {
	airfill = ( o2frac * fillpressure ) / 0.21;
	n2fill = fillpressure - airfill;
	airfillltr=airfill*tradecylindersize;
	airfillprice=airfillltr*airprice;
	sprintf(stitle, "Nitrogen fill=%.1f%s",n2fill*psifactor,porb);
	_outtext( stitle) ;
	_settextposition( row++,20);
	sprintf(stitle, "Air fill   : %.1f%s, %.1f%s, %s%.3f", airfill*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
	_outtext( stitle) ;
      }
      else {
	airfill = ( n2frac * fillpressure ) / 0.79 ;
	o2fill = fillpressure - airfill;
	o2fillltr=o2fill*tradecylindersize;
	o2fillprice=o2fillltr*o2price;
	airfillltr=airfill*tradecylindersize;
	airfillprice=airfillltr*airprice;
	sprintf(stitle, "Oxygen fill: %.1f%s, %.1f%s, %s%.3f", o2fill*psifactor, porb, o2fillltr/cuft_ltr_factor, cuftorltr, gascurrency, o2fillprice);
	_outtext( stitle) ;
	_settextposition( row++,20);
	sprintf(stitle, "Air fill   : %.1f%s, %.1f%s, %s%.3f", airfill*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
	_outtext( stitle) ;
      }
    }
  }
  if(topup) {
    airfillltr=airtopoff*tradecylindersize;
    airfillprice=airfillltr*airprice;
    _settextposition( row++,20);
    sprintf(stitle, "Air topoff: %.1f%s, %.1f%s, %s%.3f", airtopoff*psifactor, porb, airfillltr/cuft_ltr_factor, cuftorltr, gascurrency, airfillprice);
    _outtext( stitle) ;
  }
  row++;
  if( vc.numcolors > 2)_settextcolor(7);
  helpscreen(48);
  numdivers=0;
  do {
    _settextposition( row,20);
    sprintf(stitle, "Enter divers name: _______________");
    _outtext( stitle) ;
    _settextposition( row++,39);
    cnumbuf[0]=16;
    numbuf = cgetsa( cnumbuf );
    if(!*numbuf && !numdivers) {
      row++;
      _settextposition( row++,20);
      sprintf(stitle, "Press any key to continue.....");
      _outtext( stitle) ;
      getch();
      return;
    }
    strcpy( divername[numdivers], numbuf );
    if(*numbuf) numdivers++;
  } while (*numbuf);
  _strdate( tradedatebuf );
  helpscreen(49);
  _settextposition( row++,20);
  sprintf(stitle, "Save to log file?   Y/<N> ");
  _outtext( stitle) ;

  c = getchyn();
  sprintf(stitle, "%c",c);
  _outtext( stitle) ;
  if(c=='y' || c=='Y') {
    airfillcosttotal += airfillprice * (aircost/airprice) * numdivers;
    hefillcosttotal += hefillprice * (hecost/heprice) * numdivers;
    o2fillcosttotal += o2fillprice * (o2cost/o2price) * numdivers;
    airfillpricetotal += airfillprice * numdivers;
    hefillpricetotal += hefillprice * numdivers;
    o2fillpricetotal += o2fillprice * numdivers;
    savefillcosts();
    printtradecylinderfill(
    i, c, heinc, topup, entergaspercent,
    sftemp, fillpressure, workingdepth, workingppo2, o2frac, n2frac, hefrac, narcpressure, hefill, airfill, o2fill, n2fill,
    message,
    title, row,
    hefillpressure, o2fillpressure, hefillfrac, o2fillfrac, n2fillfrac, airtopoff,
    1
    );
  }
  helpscreen(39);
  _settextposition( row++,20);
  sprintf(stitle, "Send to printer?   Y/<N> ");
  _outtext( stitle) ;

  c = getchyn();
  sprintf(stitle, "%c",c);
  _outtext( stitle) ;
  if(c=='y' || c=='Y') {
    printtradecylinderfill(
    i, c, heinc, topup, entergaspercent,
    sftemp, fillpressure, workingdepth, workingppo2, o2frac, n2frac, hefrac, narcpressure, hefill, airfill, o2fill, n2fill,
    message,
    title, row,
    hefillpressure, o2fillpressure, hefillfrac, o2fillfrac, n2fillfrac, airtopoff,
    0
    );
  }
}
