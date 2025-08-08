/*=============================================================
 Project : Knowledge Management – Helpdesk Tables
 Author  : Mohammad Talal Naseem
 Purpose : Generate tables for the KMS pre/post study
 Inputs  : XLSX with sheet "Combined" (or overridable)
 Outputs : One Excel file with 4 sheets (Tables 1–4)
 Notes   : Uses synthetic/demo data when placed in /data
==============================================================*/

/*------------------ User Parameters -------------------------*/
%let in_xlsx  = ./data/KM_Data.xlsx;   /* or full path */
%let in_sheet = Combined;
%let out_dir  = ./out;
%let out_xlsx = &out_dir./KM_Data_Tables.xlsx;
%let out_rtf  = &out_dir./KM_Data_Tables.rtf;         /* optional */

/*------------------ Housekeeping ----------------------------*/
options mprint mlogic symbolgen nodate nonumber;
ods _all_ close;

/*------------------ Import ----------------------------------*/
filename reffile "&in_xlsx";
proc import datafile=reffile
    out=work.km dbms=xlsx replace;
    sheet="&in_sheet";
    getnames=yes;
run;

/*------------------ Derive N for Titles ---------------------*/
proc sql noprint;
    select count(*) into :_N trimmed from work.km;
quit;

/*================== Table 1: Category =======================*/
proc freq data=work.km order=freq noprint;
    table category / nocum out=work.table1;
run;

title "Table 1 Frequency of cases by category (N = &_N)";
proc report data=work.table1 nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black
                   borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt
                   color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white
                   borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Category count percent;
    define count   / "Frequency"   style(column)=[just=center cellwidth=1in];
    define Percent / "Percent"     format=8.2    style(column)=[just=center cellwidth=1in];
run; title;

/*================== Table 2: Priority =======================*/
proc freq data=work.km order=internal noprint;
    table priority / nocum out=work.table2;
run;

title "Table 2 Classification of cases by priority (N = &_N)";
proc report data=work.table2 nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black
                   borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt
                   color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white
                   borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Priority count percent;
    define count   / "Frequency" style(column)=[just=center cellwidth=1in];
    define Percent / "Percent"   format=8.      style(column)=[just=center cellwidth=1in];
run; title;

/*================== Table 3: Response Time by Phase & Priority =======================*/
/* Prep */
proc sort data=work.km; by priority; run;

proc means data=work.km n median qrange nway;
    class phase;
    var Response_Time_Hours;
    by priority;
    output out=work.class_means n= median= qrange= / autoname;
run;

proc means data=work.km n median qrange nway;
    class phase;
    var Response_Time_Hours;
    output out=work.class_means_all n= median= qrange= / autoname;
run;

data work.class_means_all; set work.class_means_all; Priority='All'; run;
data work.means; set work.class_means work.class_means_all; run;

/* Nonparametric test (Wilcoxon / KW as per original) */
proc npar1way data=work.km wilcoxon correct=no;
    class Phase;
    var Response_Time_Hours;
    by priority;
    output out=work.nparway;
run;

proc npar1way data=work.km wilcoxon correct=no;
    class Phase;
    var Response_Time_Hours;
    output out=work.nparway_all;
run;

data work.nparway_all; set work.nparway_all; Priority='All'; run;
data work.cox; set work.nparway work.nparway_all; run;

proc sort data=work.cox;  by Priority; run;
proc sort data=work.means; by Priority; run;

data work.table3;
    merge work.means(in=a) work.cox(in=b);
    by Priority;
    if a=1 and b=1;
run;

data work.table3(keep=Priority Phase Response_time_hours_N Response_time_hours_Median Response_time_hours_qrange P_kw);
    set work.table3;
    if phase=. then delete;
run;

/* Formats */
proc format;
    value $Pri
        'All'   = 'Total average case resolution time (hours)'
        'High'  = 'Total average time for resolving High priority cases (hours)'
        'Medium'= 'Total average time for resolving medium priority cases (hours)'
        'Low'   = 'Total average time for resolving low priority cases (hours)';
    value Phase
        1='Phase I'
        2='Phase II';
run;

title 'Table 3 Summary output for Phase I versus Phase II';
proc report data=work.table3 nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black
                   borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt
                   color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white
                   borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Priority Phase,(Response_time_hours_N Response_time_hours_Median Response_time_hours_qrange) P_kw;
    define Priority / "Summary output for Phase I versus Phase II" group format=$pri.;
    define Phase    / across format=Phase.;
    define Response_time_hours_N       / "N";
    define Response_time_hours_Median  / "Median";
    define Response_time_hours_qrange  / "IQR";
    define P_kw    / group format=pvalue. "p-value";
run; title;

/*================== Table 4: Support Level by Phase =======================*/
proc freq data=work.km noprint;
    tables Support_level*Phase / measures chisq out=work.table4 outpct;
run;

proc format;
    value SL
        1='Percentage of cases that went to 1st level'
        2='Percentage of cases that went to 2nd level'
        3='Percentage of cases that went to 3rd level';
run;

title 'Table 4 Percentages of cases referred to each support level';
proc report data=work.table4 nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black
                   borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt
                   color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white
                   borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Support_level Phase,(Count PCT_col);
    define Support_level / "Percentages of cases referred to each support" group format=SL.;
    define Phase        / across format=Phase.;
    define Count        / "N";
    define PCT_col      / "%" format=8.1;
run; title;

/*================== Output: Excel (recommended) =======================*/
ods excel file="&out_xlsx" options(embedded_titles='yes');
title "Table 1 Frequency of cases by category (N = &_N)";
proc print data=work.table1; run;

title "Table 2 Classification of cases by priority (N = &_N)";
proc print data=work.table2; run;

title "Table 3 Summary output for Phase I versus Phase II";
proc print data=work.table3; run;

title "Table 4 Percentages of cases referred to each support level";
proc print data=work.table4; run;
ods excel close; title;

/*================== Optional: RTF (matches your original) ==============*/
ods rtf file="&out_rtf" startpage=no notoc_data;
title "Table 1 Frequency of cases by category (N = &_N)";  proc report data=work.table1;  columns _all_; run;
title "Table 2 Classification of cases by priority (N = &_N)"; proc report data=work.table2; columns _all_; run;
title "Table 3 Summary output for Phase I versus Phase II"; proc report data=work.table3; columns _all_; run;
title "Table 4 Percentages of cases referred to each support level"; proc report data=work.table4; columns _all_; run;
ods rtf close; title;
