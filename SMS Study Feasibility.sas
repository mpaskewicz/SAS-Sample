*Import data pulled From epic;
FILENAME REFFILE 'H:/DM2 Study Feasibility1.xlsx';
options validvarname=v7;
PROC IMPORT DATAFILE=REFFILE
	DBMS=XLSX
	OUT=WORK.IMPORT replace;
	GETNAMES=YES;
RUN;

PROC CONTENTS DATA=WORK.IMPORT; RUN;


FILENAME REFFILE 'H:/DM2 Study Feasibility2.xlsx';
options validvarname=v7;
PROC IMPORT DATAFILE=REFFILE
	DBMS=XLSX
	OUT=WORK.IMPORT1;
	GETNAMES=YES;
RUN;

PROC CONTENTS DATA=WORK.IMPORT1; RUN;

FILENAME REFFILE 'H:/DM2 Study Feasibility3.xlsx';
options validvarname=v7;
PROC IMPORT DATAFILE=REFFILE
	DBMS=XLSX
	OUT=WORK.IMPORT2;
	GETNAMES=YES;
RUN;

PROC CONTENTS DATA=WORK.IMPORT2; RUN;

	

*Concatenate data sets and select only GIM pts where 'Type 1' does not appear in the problem list;
Data StudyFeasibility;
	set import import1 import2;
	format Last_HBA1c_Val 9.2 age 2.;
	where PCP_Department contains "PRIMARY CARE" 
		and index(Problem_list, 'type 1')=0;
*A1c value importing as character.  Correct to Numeric;
	if index(Last_HBA1c_Value, '>')>1 then Last_HBA1c_Val=14.1;
		else if index(Last_HBA1c_Value, '<')>1 then Last_HBA1c_Val=3.9;
		else Last_HBA1c_Val=Last_HBA1c_Value;
*Diabetes Med Adherence value importing as character.  Correct to Numeric;
	if Diabetes_med_adherence = "?" then DM_Med_adherence =.;
		else if Diabetes_med_adherence = "N/A" then DM_Med_adherence =.;
		else DM_Med_adherence = Diabetes_med_adherence;
*Check for dementia, schizophrenia, and down syndrome;
	if index(problem_list, 'dementia')>0 then CogDefExclusion = "Y";
		else if index(problem_list, 'down')>0 then CogDefExclusion = "Y";
		else if index(problem_list, 'schizo')>0 then CogDefExclusion = "Y";
		else CogDefExclusion = "N";
*Calculate Age;
	DOB = DOB - 21916;
	Age = yrdif(DOB, Today(), 'Age');
*Time since last A1c Measure;
	A1c6mo = intck('days', Last_HBA1c_date, today() );
*Check for at least 1 oral med;
array nn (33) $ 18 _temporary_ ("Glipizide" "Glucotrol", "Glimpepride", "Amaryl", "Glyburide", "Diabeta", "Glynase", "Metformin", "Glucophage", "Glumetza",
			"Fortamet", "Riomet", "Pioglitozone", "Rosiglitozone", "Acarbose", "Precose", "Miglitol", "Glyset", "Repaglinide", "Prandin", "Sitagliptin", 
			"Januvia", "Saxagliptin", "Onglyza", "Lingagliptin", "Trajenta", "Alogliptin", "Canagliflozin", "Invokana", "Dapagloflozin", "Farxiga", "Empagliflozin", "Jardiance");	  
	  do i=1 to dim(nn);
      if find(Current_Medications,strip(nn[i]),'i')>0 then do;
         found=1;
         match=nn[i];
         leave;
      end;
   end;
*MPR only has missing data.  Dropping this variable;
	drop MPR i;
run;

proc sort data=studyfeasibility nodupkey dupout=work.dups; by MRN;run;
%pcmh(datafile="G:\PCMH registry 2016\Registry Exports\2022-03-22_PCMH_Regsitry.xlsx", out=PCMH3_22);
data Studyfeasibility;
	merge studyfeasibility (in=a) PCMH3_22(Keep=BMC_Lab_las_hem_a1c_val BMC_Lab_las_hem_a1c_dat MRN);
	by MRN;
	if Last_HBA1c_val =. then Last_HBA1c_val = BMC_Lab_las_hem_a1c_val;
	if A1c6mo =. then a1c6m = intck('days', BMC_Lab_las_hem_a1c_dat, today() );
	if Pref_language not in("English", "Haitian Creole", "Spanish") then Pref_Language = "Other";
	if a;
run;

proc sql;
select count(distinct mrn) as Patient_Count,
	pref_language as Preferred_Language
	from studyfeasibility
	where age between 18 and 75
		and Last_HBA1c_val >=9
		and CogDefExclusion ="N"
		and match is not null
		and A1c6mo < 181
	group by Pref_Language
	order by Patient_count desc;
	
	
create table eligiblePt as
select MRN,
	pref_language as Preferred_Language,
	Last_HBA1c_val,
	DM_med_adherence,
	Pref_language,
	Mychart_Status,
	Mychart
	from studyfeasibility
	where age between 18 and 75
		and Last_HBA1c_val >=9
		and CogDefExclusion ="N"
		and match is not null
		and A1c6mo < 181;
quit;

TITLE 'Preferred Language for Eligible Patients';
ods noproctitle;
PROC GCHART DATA=EligiblePt;
      PIE Pref_language/ DISCRETE VALUE=INSIDE
                 PERCENT=INSIDE SLICE=OUTSIDE
                 plabel=(font='Albany AMT/bold' h=1.3);
                 
RUN;


proc univariate data=eligiblept;
	var Last_HBA1c_val DM_Med_adherence;
run;

proc sgplot data=eligiblept;
	title "My Chart Activation Status as Proxy for Text Message Capability";
	vbar MYchart_status;
run;
	
