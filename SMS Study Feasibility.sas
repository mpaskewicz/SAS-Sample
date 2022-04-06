*Import data pulled From epic;
%macro Importdata(filepath, out);
options validvarname=v7;
PROC IMPORT DATAFILE=&filepath
	DBMS=XLSX
	OUT=WORK.&out replace;
	GETNAMES=YES;
RUN;

PROC CONTENTS DATA=WORK.&out; RUN;
%mend
%importdata(filepath="...\examplefile1.xlsx", out=Import1);
%importdata(filepath="...\examplefile2.xlsx", out=Import2);
%importdata(filepath="...\examplefile3.xlsx", out=Import3);
	

*Concatenate data sets and select pts from my department that are not type 1 diabetics.;
Data StudyFeasibility;
	set import1 import2 import3;
	format Last_HBA1c_Val 9.2 age 2.;
	where PCP_Department contains "My Dept" 
		and index(Problem_list, 'type 1')=0;
*A1c value importing as character because of values like '>14'.  Correct to Numeric;
	if index(Last_HBA1c_Value, '>')>1 then Last_HBA1c_Val=14.1;
		else if index(Last_HBA1c_Value, '<')>1 then Last_HBA1c_Val=3.9;
		else Last_HBA1c_Val=Last_HBA1c_Value;
*Diabetes Med Adherence value importing as character due to letter/special characters.  Correct to Numeric;
	if Diabetes_med_adherence = "?" then DM_Med_adherence =.;
		else if Diabetes_med_adherence = "N/A" then DM_Med_adherence =.;
		else DM_Med_adherence = Diabetes_med_adherence;
*Check for exclusionary criteria;
	if index(problem_list, 'dementia')>0 then CogDefExclusion = "Y";
		else if index(problem_list, 'down')>0 then CogDefExclusion = "Y";
		else if index(problem_list, 'schizo')>0 then CogDefExclusion = "Y";
		else CogDefExclusion = "N";
*Calculate from Excel Datetime;
	DOB = DOB - 21916;
	Age = yrdif(DOB, Today(), 'Age');
*Time since last A1c Measure;
	A1c6mock = intck('days', Last_HBA1c_date, today() );
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
*prxmatch could also be used to search for specific strings within current medications;
*MPR only has missing data.  Dropping this variable;
	drop MPR i;
run;
proc sort data=studyfeasibility nodupkey dupout=work.dups; by MRN;run;
*Run local macro import registry data then merge with Epic data.  Supplemnt missing data where necessary in the data step.  A proc sql join with the 
Coalesce fucntion could acheive the same result;
%pcmh(datafile=".../2022-03-22_PCMH_Regsitry.xlsx", out=PCMH3_22);
data Studyfeasibility;
	merge studyfeasibility (in=a) PCMH3_22(Keep=BMC_Lab_las_hem_a1c_val BMC_Lab_las_hem_a1c_dat MRN);
	by MRN;
	if Last_HBA1c_val =. then Last_HBA1c_val = BMC_Lab_las_hem_a1c_val;
	if A1c6mo_Ck =. then a1c6mo_Ck = intck('days', BMC_Lab_las_hem_a1c_dat, today() );
	if Pref_language not in("English", "Haitian Creole", "Spanish") then Pref_Language = "Other";
	if a;
run;

proc sql;
Title "Overall Number of Eligle Patients";
select count(mrn) as Patient_Count,
	pref_language as Preferred_Language
	from studyfeasibility
	where age between 18 and 75
		and Last_HBA1c_val >=9
		and CogDefExclusion ="N"
		and match is not null
		and A1c6moCK < 181
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
		and A1c6mock < 181;
quit;

TITLE 'Preferred Language for Eligible Patients';
ods noproctitle;
PROC GCHART DATA=EligiblePt;
      PIE Pref_language/ DISCRETE VALUE=INSIDE
                 PERCENT=INSIDE SLICE=OUTSIDE
                 plabel=(font='Albany AMT/bold' h=1.3);
                 
RUN;
title;

*Assess Diabetic med adherence data with descriptive statistics;
proc univariate data=eligiblept;
	var DM_Med_adherence;
run;

proc sgplot data=eligiblept;
	title "My Chart Activation Status as Proxy for Ability to recieve SMS messages";
	vbar MYchart_status;
run;
	
