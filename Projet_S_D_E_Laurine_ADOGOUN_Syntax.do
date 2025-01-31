********************************************************************************
*                                                                              *
*           Program Evaluation for International Development                   *
*                    Replication Project 2024-2025                             *
*Using Maimonides' Rule to Estimate the Effect of Class Size on Scholastic     *
*     Achievement Author(s): Joshua D. Angrist and Victor Lavy                 *
*                                                                              *
********************************************************************************

***Preparation of work environment***
global workdir "C:\Users\DELL\Desktop\ISE3\Semestre1\Evaluation_d'impacts_des_politiques_publiques\Projet_S_D_E_Laurine_ADOGOUN"

cd "${workdir}"

global Data "${workdir}\Data"

global Outputs "${workdir}\Outputs"


***preliminary stuff*** 
clear all  
set scheme s1mono  
set more off 
set seed 12345 
version 16.1


***Data combining***
use "${Data}\Grade 4.dta", clear
append using "${Data}\Grade 5.dta"

save "${Data}\grades_combined.dta", replace

***Rename some variables***
rename (classize c_size tipuach verbsize mathsize avgverb avgmath) (class_size enrollment percent_disadvantaged reading_size math_size average_verbal average_math)

***Pre-processing***
replace average_verbal= average_verbal-100 if average_verbal>100
replace average_math= average_math-100 if average_math>100

replace average_verbal=. if reading_size==0
replace average_math=. if math_size==0

count if average_math==. & average_verbal==. & grade==5
count if average_math==. & average_verbal==. & grade==4

drop if average_math==. & average_verbal==.

/* The empirical analysis is restricted to schools with at least 5 pupils
reported enrolled in the relevant grade and to classes with less than 45 pupils.*/
keep if 1<class_size & class_size<45 & enrollment>5


***Create a label to identify grades***
gen grade_label = "5th grade" if grade == 5
replace grade_label = "4th grade" if grade == 4

********************************************************************************
******************             Table1                   ************************
********************************************************************************

* Compute descriptive statistics for each grade
preserve

* Compute descriptive statistics for 4th grade
tabstat class_size enrollment percent_disadvantaged reading_size math_size average_verbal average_math , statistics(mean sd p10 p25 p50 p75 p90) format(%9.1f) long by(grade_label) columns(statistics)
restore		
********************Export the final table to Excel*****************************

***Define the Excel output path***
putexcel set "${Outputs}\Table1.xlsx", replace

***Create a binary variable for unique schools***
gen is_unique_school = schlcode != schlcode[_n-1]  // Directly assigns 1 for each unique school
sort schlcode  

* Create a binary variable for unique classes
gen is_unique_class = class_size != .  // Assign 1 for each class with a valid size (non-missing)

***Write column headers in Excel***
putexcel A1="Grade" B1="Variable" C1="Mean" D1="S.D." E1="0.10" F1="0.25" ///
         G1="0.50" H1="0.75" I1="0.90", bold
		 	 
		 
					**********Part 1: Full sample***********
					
					
putexcel A2 = "A. Full sample", bold

***Define the row number where the data starts***
local row = 3  

***Define variables to summarize***
local variables class_size enrollment percent_disadvantaged reading_size math_size average_verbal average_math

***Loop through both grades***
foreach g in "5th grade" "4th grade" {
	
    * Add the grade title in Excel
	summarize is_unique_school  if grade_label=="`g'"
	local num_schools = r(sum)
	summarize is_unique_class if grade_label=="`g'"
	local num_classes = r(sum) 
    putexcel A`row' = "`g'(`num_classes' classes, `num_schools' schools, tested in 1991)", bold
    local row = `row' + 1
    
    * Compute descriptive statistics for the selected grade
    foreach var in `variables' {
        summarize `var' if grade_label == "`g'", detail
        
        * Store results in local macros
        local Mean = round(r(mean), 0.1)
        local SD = round(r(sd), 0.1)

        * Compute percentiles
        centile `var', centile(10 25 50 75 90)
        local p10 = round(r(c_1), 0.1)
        local p25 = round(r(c_2), 0.1)
        local p50 = round(r(c_3), 0.1)
        local p75 = round(r(c_4), 0.1)
        local p90 = round(r(c_5), 0.1)

        * Export results to Excel
        putexcel B`row' = "`var'" C`row' = `Mean' D`row' = `SD' E`row' = `p10' F`row' = `p25' ///
                 G`row' = `p50' H`row' = `p75' I`row' = `p90'

        * Move to the next row
        local row = `row' + 1
    }
}


					**********Part 2: Discontinuity***********
					
***Creation of the Excel file to export results***
putexcel set "$Outputs\Table1.xlsx", modify 

***Add the section title in Excel***
putexcel A`row' = ("B. +/- 5 Discontinuity sample (enrollment 36-45, 76-85, 116-124)"), bold

***Create a binary Enrollment variable***
gen Enrollment = 0  // Initialize the Enrollment variable to 0
replace Enrollment = 1 if inrange(enrollment, 36, 45) | inrange(enrollment, 76, 85) | inrange(enrollment, 116, 125)  // Set Enrollment to 1 for the specified ranges

***Create a binary variable for unique classes***
gen is_unique_class_dis = 0  
sort classid  
replace is_unique_class_dis = 1 if (Enrollment == 1) & (class_size != .)  // Set to 1 for each unique class (non-missing class size)


* Add grades and indicators
local row = `row' + 1  // Increment the row counter
putexcel (C`row':D`row') ="5th grade", hcenter bold border(top bottom) merge  // Add Grade 5 header with centering, bold, and borders
putexcel (E`row':F`row') ="4th grade", hcenter bold border(top bottom) merge  

local row = `row' + 1  // Increment the row counter for the next line
putexcel C`row'=("Mean") D`row'=("S.D.") E`row'=("Mean") F`row'=("S.D."), hcenter bold  

local start_row = `row'  // Store the current row for future reference
local row = `row' + 1


***Discontinuity for 5th grade***	
***Create a binary variable for unique schools***
preserve
keep if grade_label=="5th grade"
gen is_unique_school_dis5 = 0  
sort schlcode 
replace is_unique_school_dis5 = 1 if (Enrollment == 1) & (schlcode != schlcode[_n-1])  // Set to 1 for each unique school (compared to the previous school)

* Add the grade title in Excel
	summarize is_unique_school_dis5  
	local num_schools_dis5 = r(sum)
	summarize is_unique_class_dis 
	local num_classes_dis5 = r(sum) 
    putexcel (C`row':D`row')= "(`num_classes_dis5' classes, `num_schools_dis5' schools)", hcenter bold merge
    local row = `row' + 1
    
    * Compute descriptive statistics for the selected grade
    foreach var in `variables' {
        summarize `var' , detail
        
        * Store results in local macros
        local Mean = round(r(mean), 0.1)
        local SD = round(r(sd), 0.1)

        * Export results to Excel
        putexcel B`row' = "`var'" C`row' = `Mean' D`row' = `SD'

        * Move to the next row
        local row = `row' + 1
    }
restore

***Discontinuity for 4th grade***	
preserve
keep if grade_label=="4th grade"
gen is_unique_school_dis4 = 0  
sort schlcode 
replace is_unique_school_dis4 = 1 if (Enrollment == 1) & (schlcode != schlcode[_n-1])   

* Add the grade title in Excel
local row = `start_row' 
local row = `row' + 1
	summarize is_unique_school_dis4  
	local num_schools_dis4 = r(sum)
	summarize is_unique_class_dis 
	local num_classes_dis4 = r(sum) 
    putexcel (E`row':F`row')= "(`num_classes_dis4' classes, `num_schools_dis4' schools)", hcenter bold merge
    local row = `row' + 1
    
    * Compute descriptive statistics for the selected grade
    foreach var in `variables' {
        summarize `var', detail
        
        * Store results in local macros
        local Mean = round(r(mean), 0.1)
        local SD = round(r(sd), 0.1)

        * Export results to Excel
        putexcel E`row' = `Mean' F`row' = `SD'

        * Move to the next row
        local row = `row' + 1
    }
restore

*** Final message ***
di "Table 1 successfully exported to Excel!"



********************************************************************************
******************             Figure2                  ************************
********************************************************************************

					**********a: Fifth grade***********
preserve
keep if grade_label=="5th grade"

***Create the interval variable to group enrollments in intervals of 10***
gen interval = floor((enrollment - 1) / 10) * 10  // Group enrollment into intervals of 10
replace interval = 160 if interval > 160  // Set a cap for the interval at 160

***Calculate the midpoint for each enrollment interval***
gen midpoint = interval + 5  // Midpoint is the center of each interval

***Calculate predicted class size***
gen pred_class = ceil(enrollment / 40)  // Round up to the nearest whole class size for each enrollment
gen pred_size = enrollment / pred_class  // Calculate the predicted class size

***Filter the data to include only grade 5 students with enrollment between 10 and 190***
keep if (enrollment >= 10 & enrollment <= 190)

***Calculate the mean reading score and predicted class size for each interval***
bysort interval (midpoint): egen mean_verb = mean(average_verbal)  
bysort interval (midpoint): egen pred_size_mean = mean(pred_size) 

***Create a graph for the fifth grade***
twoway /// 
    (line mean_verb midpoint, lcolor(blue) lwidth(medium) lpattern(solid)) /// 
    (line pred_size_mean midpoint, lcolor(purple) lwidth(medium) lpattern(dash) yaxis(2)), /// 
    ytitle("Average reading score", axis(1)) /// 
    ytitle("Average size function", axis(2)) ///  
    ylabel(70(1)80, axis(1) angle(0)) ///  
    ylabel(5(5)40, axis(2) angle(0)) ///  
    xlabel(5(20)165) ///  
    xscale(range(5 165)) ///  
    xtitle("Enrollment count") ///  
    yscale(range(70 80) axis(1)) /// 
    yscale(range(5 .) axis(2)) ///  
    yline(20.5 27 30.25 32.2 33.33 40, lpattern(shortdash) lwidth(0.15) lcolor(black) axis(2)) /// 
    title("a. Fifth Grade") ///  
    legend(off) /// 
    text(79 45 "Predicted class size", color(purple) size(medium)) ///  
    text(73 80 "Average test scores", color(blue) size(medium)) ///  
    name(fifth_grade, replace) 
graph export "$Outputs\Fifth_Graph(2_a).png", replace width(1500) height(900)  // Export the graph as a PNG image
restore


					**********b: Fourth grade***********
preserve
keep if grade_label=="4th grade"

***Create the interval variable to group enrollments in intervals of 10***
gen interval = floor((enrollment - 1) / 10) * 10  
replace interval = 160 if interval > 160  

***Calculate the midpoint for each enrollment interval***
gen midpoint = interval + 5  

***Calculate predicted class size***
gen pred_class = ceil(enrollment / 40)  
gen pred_size = enrollment / pred_class  

***Filter the data to include only grade 5 students with enrollment between 10 and 190***
keep if (enrollment >= 10 & enrollment <= 190) & (midpoint > 5)

***Calculate the mean reading score and predicted class size for each interval***
bysort interval (midpoint): egen mean_verb = mean(average_verbal)  
bysort interval (midpoint): egen pred_size_mean = mean(pred_size) 

***Create a graph for the fifth grade***
twoway ///
    (line mean_verb midpoint, lcolor(blue) lwidth(medium) lpattern(solid)) ///
    (line pred_size_mean midpoint, lcolor(purple) lwidth(medium) lpattern(dash) yaxis(2)), ///
    ytitle("Average reading score", axis(1)) ///
    ytitle("Average size function", axis(2)) ///
    ylabel(68(1)78, axis(1) angle(0)) ///
    ylabel(5(5)40, axis(2) angle(0)) ///
    xlabel(5(20)165) ///
    xscale(range(5 165)) ///
    xtitle("Enrollment count") ///
    yscale(range(68 78) axis(1)) ///
    yscale(range(5 .) axis(2)) ///
    yline(72.4 74.3 75.2 75.8 76.2 78, lpattern(shortdash) lwidth(0.15) lcolor(black) axis(1)) ///
    title("b. Fourth Grade") ///
    legend(off) /// 
    text(77 45 "Predicted class size", color(purple) size(small)) /// 
    text(71.4 80 "Average test scores", color(blue) size(small)) /// 
    name(fifth_grade, replace)
graph export "$Outputs\Fourth_Graph(2_b).png", replace width(1500) height(900)  
restore





********************************************************************************
******************             Figure3                  ************************
********************************************************************************
					**********a: Fifth grade (Reading)***********
preserve
keep if grade_label=="5th grade"

***Create the interval variable to group enrollments in intervals of 10***
gen interval = floor((enrollment - 1) / 10) * 10  // Group enrollment into intervals of 10
replace interval = 160 if interval > 160  // Set a cap for the interval at 160

***Calculate the midpoint for each enrollment interval***
gen midpoint = interval + 5  // Midpoint is the center of each interval

***Calculate predicted class size***
gen pred_class = ceil(enrollment / 40)  // Round up to the nearest whole class size for each enrollment
gen pred_size = enrollment / pred_class  // Calculate the predicted class size


***Filter the data to include only grade 5 students with enrollment between 10 and 190***
keep if (enrollment >= 9 & enrollment <= 190)

***Regressions of average reading scores and the average of predicted size on average enrollment and PD index for each interval
regress average_verbal enrollment percent_disadvantaged, cluster(schlcode)
predict residuals1, resid

regress pred_size enrollment percent_disadvantaged, cluster(schlcode)
predict residuals2, resid

***Calculate the mean reading score and predicted class size for each interval***
bysort interval (midpoint): egen res_mean_verb = mean(residuals1)  
bysort interval (midpoint): egen res_pred_size_mean = mean(residuals2) 

keep if midpoint>5

***Create a graph for the fifth grade***
twoway /// 
    (line res_mean_verb midpoint, lcolor(blue) lwidth(medium) lpattern(solid)) /// 
    (line res_pred_size_mean midpoint, lcolor(purple) lwidth(medium) lpattern(dash) yaxis(2)), /// 
    ytitle("Reading score residual", axis(1)) /// 
    ytitle("Size-function residual", axis(2)) ///  
    ylabel(-5(1)5, axis(1) angle(0)) ///  
    ylabel(-15(5)15, axis(2) angle(0)) ///  
    xlabel(5(20)165) ///  
    xscale(range(5 165)) ///  
    xtitle("Enrollment count") ///  
    yscale(range(-5 5) axis(1)) /// 
    yscale(range(-15 15) axis(2)) ///  
    title("a. Fifth Grade (Reading)") ///  
    legend(off) /// 
    text(4 45 "Predicted class size", color(purple) size(medium)) ///  
    text(-3.5 80 "Average test scores", color(blue) size(medium)) ///  
    name(fifth_grade, replace) 
graph export "$Outputs\Fifth_Graph(3_a).png", replace width(1800) height(500)
restore

								
					
					**********b: Fourth grade (Reading)***********
preserve
keep if grade_label=="4th grade"

***Create the interval variable to group enrollments in intervals of 10***
gen interval = floor((enrollment - 1) / 10) * 10  
replace interval = 160 if interval > 160  

***Calculate the midpoint for each enrollment interval***
gen midpoint = interval + 5  

***Calculate predicted class size***
gen pred_class = ceil(enrollment / 40)  
gen pred_size = enrollment / pred_class  



***Filter the data to include only grade 5 students with enrollment between 10 and 190***
keep if (enrollment >= 9 & enrollment <= 190)

***Regressions of average reading scores and the average of predicted size on average enrollment and PD index for each interval
regress average_verbal enrollment percent_disadvantaged
predict residuals1, resid

regress pred_size enrollment percent_disadvantaged
predict residuals2, resid

***Calculate the mean reading score and predicted class size for each interval***
bysort interval (midpoint): egen res_mean_verb = mean(residuals1) 
bysort interval (midpoint): egen res_pred_size_mean = mean(residuals2) 

keep if midpoint>5

***Create a graph for the fifth grade***
twoway /// 
    (line res_mean_verb midpoint, lcolor(blue) lwidth(medium) lpattern(solid)) /// 
    (line res_pred_size_mean midpoint, lcolor(purple) lwidth(medium) lpattern(dash) yaxis(2)), /// 
    ytitle("Reading score residual", axis(1)) /// 
    ytitle("Size-function residual", axis(2)) ///  
    ylabel(-5(1)5, axis(1) angle(0)) ///  
    ylabel(-15(5)15, axis(2) angle(0)) ///  
    xlabel(5(20)165) ///  
    xscale(range(5 165)) ///  
    xtitle("Enrollment count") ///  
    yscale(range(-5 5) axis(1)) /// 
    yscale(range(-15 15) axis(2)) ///  
    title("a. Fourth Grade (Reading)") ///  
    legend(off) /// 
    text(4 45 "Predicted class size", color(purple) size(medium)) ///  
    text(-3.5 80 "Average test scores", color(blue) size(medium)) ///  
    name(fifth_grade, replace) 
graph export "$Outputs\Fourth_Graph(3_b).png", replace width(1800) height(500)
restore

					
					
					
					**********c: Fifth grade (Math)***********
preserve
keep if grade_label=="5th grade"

***Create the interval variable to group enrollments in intervals of 10***
gen interval = floor((enrollment - 1) / 10) * 10  
replace interval = 160 if interval > 160  

***Calculate the midpoint for each enrollment interval***
gen midpoint = interval + 5  

***Calculate predicted class size***
gen pred_class = ceil(enrollment / 40)  
gen pred_size = enrollment / pred_class  


***Filter the data to include only grade 5 students with enrollment between 10 and 190***
keep if (enrollment >= 9 & enrollment <= 190)

***Regressions of average reading scores and the average of predicted size on average enrollment and PD index for each interval
regress average_math enrollment percent_disadvantaged
predict residuals1, resid

regress pred_size enrollment percent_disadvantaged
predict residuals2, resid


***Calculate the mean reading score and predicted class size for each interval***
bysort interval (midpoint): egen res_mean_math = mean(residuals1)  
bysort interval (midpoint): egen res_pred_size_mean = mean(residuals2) 

keep if midpoint>5

***Create a graph for the fifth grade***
twoway /// 
    (line res_mean_math midpoint, lcolor(blue) lwidth(medium) lpattern(solid)) /// 
    (line res_pred_size_mean midpoint, lcolor(purple) lwidth(medium) lpattern(dash) yaxis(2)), /// 
    ytitle("Math score residual", axis(1)) /// 
    ytitle("Size-function residual", axis(2)) ///  
    ylabel(-5(1)5, axis(1) angle(0)) ///  
    ylabel(-15(5)15, axis(2) angle(0)) ///  
    xlabel(5(20)165) ///  
    xscale(range(5 165)) ///  
    xtitle("Enrollment count") ///  
    yscale(range(-5 5) axis(1)) /// 
    yscale(range(-15 15) axis(2)) ///  
    title("a. Fifth Grade (Math)") ///  
    legend(off) /// 
    text(4 45 "Predicted class size", color(purple) size(medium)) ///  
    text(-3.5 80 "Average test scores", color(blue) size(medium)) ///  
    name(fifth_grade, replace) 
graph export "$Outputs\Fifth_Graph(3_c).png", replace width(1800) height(500)
restore				
					
					
					

********************************************************************************
******************             Table2                   ************************
********************************************************************************
*** Define the Excel output path ***
putexcel set "${Outputs}\Table2.xlsx", replace

*** Header formatting ***
putexcel C2:N2 = "OLS ESTIMATES FOR 1991", merge hcenter bold
putexcel C3:H3 = "5th Grade", merge hcenter bold
putexcel I3:N3 = "4th Grade", merge hcenter bold
putexcel C4:E4 = "Reading Comprehension", merge hcenter bold
putexcel I4:K4 = "Reading Comprehension", merge hcenter bold
putexcel F4:H4 = "Math", merge hcenter bold
putexcel L4:N4 = "Math", merge hcenter bold

*** Column numbers (1 to 12) in row 5 ***
forvalues i = 1/12 {
    local value `i'
    local formatted_value (`value')
    local col `=char(64 + `i' + 2)'
    putexcel `col'5 = "`formatted_value'"
}


*** Define row labels ***
putexcel A6:B6 = "Mean score", merge hcenter bold
putexcel A7:B7 = "(s.d.)", merge hcenter bold
putexcel A8 = "Regressors", hcenter bold
putexcel B9 = "Class size", hcenter bold
putexcel B11 = "Percent disadvantaged", hcenter bold
putexcel B13 = "Enrollment", hcenter bold
putexcel B15 = "Root MSE", hcenter bold
putexcel B16 = "R²", hcenter bold
putexcel B17 = "N", hcenter bold

*** Regression loop configuration ***
local grades "5 4"   // Grades 5 and 4
local outcomes "average_verbal average_math"   // Dependent variables
local regressors "class_size percent_disadvantaged enrollment"  // Independent variables
local columns_5 "C D E F G H"  // Columns for 5th grade
local columns_4 "I J K L M N"  // Columns for 4th grade


* Loop over grades
foreach grade of local grades {
    local cols = cond(`grade' == 5, "`columns_5'", "`columns_4'")
    
    * Loop over models
    forval model = 1/3 {
        local col_verbal : word `model' of `cols'
        local col_math   : word `=`model' + 3' of `cols'

        * Loop over outcomes (verbal_score and math_score)
        foreach outcome of local outcomes {
            local col = cond("`outcome'" == "average_verbal", "`col_verbal'", "`col_math'")
            local reg_controls = ""
            
            // Construct explanatory variables based on the model
            forval i = 1/`model' {
                local reg_controls `reg_controls' `: word `i' of `regressors''  
            }
			
            summ `outcome' if grade == `grade', detail
			sca x = round(r(mean), .1)
			putexcel `col'6 = x, hcenter
			sca x = round(r(sd), .1)
			putexcel `col'7 = x, hcenter

            
			* Run the regression
            regress `outcome' `reg_controls' if grade == `grade', cluster(schlcode)

            // Store results in Excel
            local start_row = 9
            foreach var of local reg_controls {
                sca x = round(_b[`var'],.001)
                putexcel `col'`start_row' = x, hcenter
                local j = round(_se[`var'],.001)
                putexcel `col'`=`start_row' + 1' = "( `j' )", hcenter
                local start_row = `start_row' + 2
            }

            // Store statistics: RMSE, R², and sample size
            sca x = round(e(rmse),.01)
            putexcel `col'15 = x, hcenter
            sca x = round(e(r2),.001)
            putexcel `col'16 = x, hcenter
            sca x = e(N)
            putexcel `col'17 = x, hcenter
        }
    }
}

putexcel set "$Outputs\Table2.xlsx", modify 

putexcel C6:E6, merge hcenter 
putexcel C7:E7, merge hcenter 
putexcel F6:H6, merge hcenter 
putexcel F7:H7, merge hcenter 
putexcel I6:K6, merge hcenter 
putexcel I7:K7, merge hcenter 
putexcel L6:N6, merge hcenter 
putexcel L7:N7, merge hcenter 
putexcel C17:E17, merge hcenter 
putexcel F17:H17, merge hcenter 
putexcel I17:K17, merge hcenter 
putexcel L17:N17, merge hcenter 


*** Final message ***
di "Table 2 successfully exported to Excel!"


********************************************************************************
******************             Table3                   ************************
********************************************************************************
preserve

*** Define the Excel output path ***
putexcel set "${Outputs}\Table3.xlsx", replace

					**********Part 1: A.Full sample***********
putexcel A6 = "A. Full sample", bold

*** Header formatting ***
putexcel C2:N2 = "REDUCED-FORM ESTIMATES FOR 1991", merge hcenter bold
putexcel C3:H3 = "5th Grade", merge hcenter bold
putexcel I3:N3 = "4th Grade", merge hcenter bold
putexcel C4:D4 = "Class size", merge hcenter bold
putexcel I4:J4 = "Class size", merge hcenter bold
putexcel E4:F4 = "Reading Comprehension", merge hcenter bold
putexcel K4:L4 = "Reading Comprehension", merge hcenter bold
putexcel G4:H4 = "Math", merge hcenter bold
putexcel M4:N4 = "Math", merge hcenter bold

*** Column numbers (1 to 12) in row 5 ***
forvalues i = 1/12 {
    local value `i'
    local formatted_value (`value')
    local col `=char(64 + `i' + 2)'
    putexcel `col'5 = "`formatted_value'"
}

*** Define row labels ***
putexcel A7:B7 = "Means", merge hcenter bold
putexcel A8:B8 = "(s.d.)", merge hcenter bold
putexcel A9 = "Regressors", hcenter bold
putexcel B10 = "f_sc", hcenter bold
putexcel B12 = "Percent disadvantaged", hcenter bold
putexcel B14 = "Enrollment", hcenter bold
putexcel B16 = "Root MSE", hcenter bold
putexcel B17 = "R²", hcenter bold
putexcel B18 = "N", hcenter bold

*** Generate predicted class size based on Maimonides' rule ***
gen pred_class = ceil(enrollment / 40)  
gen pred_size = enrollment / pred_class 

*** Regression loop configuration ***
local grades "5 4"
local outcomes "class_size average_verbal average_math"   // Three dependent variables
local regressors_1 "pred_size percent_disadvantaged"  // Model 1 regressors
local regressors_2 "pred_size percent_disadvantaged enrollment"  // Model 2 regressors
local columns_5 "C D E F G H"  // Columns for 5th grade
local columns_4 "I J K L M N"  // Columns for 4th grade

*** Loop over grades ***
foreach grade of local grades {
    local cols = cond(`grade' == 5, "`columns_5'", "`columns_4'")

    *** Loop over models (1 = 2 regressors, 2 = 3 regressors) ***
    forval model = 1/2 {
        local col_class : word `model' of `cols'    
        local col_verbal : word `=`model' + 2' of `cols'  
        local col_math   : word `=`model' + 4' of `cols'  

        *** Loop over outcomes (class_size, average_verbal, average_math) ***
        foreach outcome of local outcomes {
            if ("`outcome'" == "class_size") {
                local col `col_class'
            }
            else if ("`outcome'" == "average_verbal") {
                local col `col_verbal'
            }
            else {
                local col `col_math'
            }

            *** Select explanatory variables based on the model ***
            if `model' == 1 {
                local reg_controls "`regressors_1'"
            }
            else {
                local reg_controls "`regressors_2'"
            }

            *** Compute summary statistics ***
            summ `outcome' if grade == `grade', detail
            sca x = round(r(mean), 0.1)
            putexcel `col'7 = x, hcenter
            sca x = round(r(sd), 0.1)
            putexcel `col'8 = x, hcenter

            *** Run the regression ***
            regress `outcome' `reg_controls' if grade == `grade', cluster(schlcode)

            *** Store regression results ***
            local start_row = 10
            foreach var of local reg_controls {
                sca x = round(_b[`var'], 0.001)
                putexcel `col'`start_row' = x, hcenter
                local j = round(_se[`var'], 0.001)
                putexcel `col'`=`start_row' + 1' = "( `j' )", hcenter
                local start_row = `start_row' + 2
            }

            *** Store model statistics ***
            sca x = round(e(rmse), 0.01)
            putexcel `col'16 = x, hcenter
            sca x = round(e(r2), 0.01)
            putexcel `col'17 = x, hcenter
            sca x = e(N)
            putexcel `col'18 = x, hcenter
        }
    }
}

					**********Part 2: Discontinuity***********
	
***Creation of the Excel file to export results***
putexcel set "$Outputs\Table3.xlsx", modify 

***Add the section title in Excel***
putexcel A19 = ("B. Discontinuity sample "), bold

keep if Enrollment==1

*** Define row labels ***
putexcel A20:B20 = "Means", merge hcenter bold
putexcel A21:B21 = "(s.d.)", merge hcenter bold
putexcel A22 = "Regressors", hcenter bold
putexcel B23 = "f_sc", hcenter bold
putexcel B25 = "Percent disadvantaged", hcenter bold
putexcel B27 = "Enrollment", hcenter bold
putexcel B29 = "Root MSE", hcenter bold
putexcel B30 = "R²", hcenter bold
putexcel B31 = "N", hcenter bold


*** Regression loop configuration ***
local grades "5 4"
local outcomes "class_size average_verbal average_math"   // Three dependent variables
local regressors_1 "pred_size percent_disadvantaged"  // Model 1 regressors
local regressors_2 "pred_size percent_disadvantaged enrollment"  // Model 2 regressors
local columns_5 "C D E F G H"  // Columns for 5th grade
local columns_4 "I J K L M N"  // Columns for 4th grade

*** Loop over grades ***
foreach grade of local grades {
    local cols = cond(`grade' == 5, "`columns_5'", "`columns_4'")

    *** Loop over models (1 = 2 regressors, 2 = 3 regressors) ***
    forval model = 1/2 {
        local col_class : word `model' of `cols'    
        local col_verbal : word `=`model' + 2' of `cols'  
        local col_math   : word `=`model' + 4' of `cols'  

        *** Loop over outcomes (class_size, average_verbal, average_math) ***
        foreach outcome of local outcomes {
            if ("`outcome'" == "class_size") {
                local col `col_class'
            }
            else if ("`outcome'" == "average_verbal") {
                local col `col_verbal'
            }
            else {
                local col `col_math'
            }

            *** Select explanatory variables based on the model ***
            if `model' == 1 {
                local reg_controls "`regressors_1'"
            }
            else {
                local reg_controls "`regressors_2'"
            }

            *** Compute summary statistics ***
            summ `outcome' if grade == `grade', detail
            sca x = round(r(mean), 0.1)
            putexcel `col'20 = x, hcenter
            sca x = round(r(sd), 0.1)
            putexcel `col'21 = x, hcenter

            *** Run the regression ***
            regress `outcome' `reg_controls' if grade == `grade', cluster(schlcode)

            *** Store regression results ***
            local start_row = 23
            foreach var of local reg_controls {
                sca x = round(_b[`var'], 0.001)
                putexcel `col'`start_row' = x, hcenter
                local j = round(_se[`var'], 0.001)
                putexcel `col'`=`start_row' + 1' = "( `j' )", hcenter
                local start_row = `start_row' + 2
            }

            *** Store model statistics ***
            sca x = round(e(rmse), 0.01)
            putexcel `col'29 = x, hcenter
            sca x = round(e(r2), 0.01)
            putexcel `col'30 = x, hcenter
            sca x = e(N)
            putexcel `col'31 = x, hcenter
        }
    }
}


putexcel set "$Outputs\Table3.xlsx", modify 

* Full sample
putexcel C7:D7, merge hcenter 
putexcel C8:D8, merge hcenter 
putexcel E7:F7, merge hcenter 
putexcel E8:F8, merge hcenter 
putexcel G7:H7, merge hcenter 
putexcel G8:H8, merge hcenter 
putexcel I7:J7, merge hcenter 
putexcel I8:J8, merge hcenter 
putexcel K7:L7, merge hcenter 
putexcel K8:L8, merge hcenter 
putexcel M7:N7, merge hcenter 
putexcel M8:N8, merge hcenter 
putexcel C18:D18, merge hcenter 
putexcel E18:F18, merge hcenter 
putexcel G18:H18, merge hcenter
putexcel I18:J18, merge hcenter 
putexcel K18:L18, merge hcenter 
putexcel M18:N18, merge hcenter


* Discontinuty sample 
putexcel C20:D20, merge hcenter 
putexcel C31:D31, merge hcenter 
putexcel C21:D21, merge hcenter 
putexcel E20:F20, merge hcenter 
putexcel E21:F21, merge hcenter 
putexcel E31:F31, merge hcenter 
putexcel G20:H20, merge hcenter 
putexcel G21:H21, merge hcenter
putexcel G31:H31, merge hcenter
putexcel I20:J20, merge hcenter 
putexcel I21:J21, merge hcenter 
putexcel I31:J31, merge hcenter 
putexcel K20:L20, merge hcenter 
putexcel K21:L21, merge hcenter 
putexcel K31:L31, merge hcenter 
putexcel M20:N20, merge hcenter 
putexcel M21:N21, merge hcenter
putexcel M31:N31, merge hcenter

*** Final message ***
di "Table 3 successfully exported to Excel!"
restore



********************************************************************************
******************             Table4                   ************************
********************************************************************************
preserve
keep if grade==5
*** Define the Excel output path ***
putexcel set "${Outputs}\Table4.xlsx", replace

*** Generate predicted class size based on Maimonides' rule ***
gen pred_class = ceil(enrollment / 40)  
gen pred_size = enrollment / pred_class 

*** Generate Enrollment squared/100 ***
gen enrollement_squared = (enrollment*enrollment)/100

*** Generate Piecewise linear trend ***
gen trend = enrollment if enrollment>=0 & enrollment<=40
replace trend= 20+(enrollment/2) if enrollment>=41 & enrollment<=80
replace trend= (100/3)+(enrollment/3) if enrollment>=81 & enrollment<=120
replace trend= (130/3)+(enrollment/4) if enrollment>=121 & enrollment<=160


*** Header formatting ***
putexcel C2:N2 = "2SLS ESTIMATES FOR 1991 (FIFTH GRADERS)", merge hcenter bold
putexcel C3:H3 = "Reading comprehension", merge hcenter bold
putexcel I3:N3 = "Math", merge hcenter bold
putexcel C4:F4 = "Full sample", merge hcenter bold
putexcel G4:H4 = "+/-5 Discontinuity sample", merge hcenter bold
putexcel I4:L4 = "Full sample", merge hcenter bold
putexcel M4:N4 = "+/-5 Discontinuity sample", merge hcenter bold

*** Column numbers (1 to 12) in row 5 ***
forvalues i = 1/12 {
    local value `i'
    local formatted_value (`value')
    local col `=char(64 + `i' + 2)'
    putexcel `col'5 = "`formatted_value'"
}

*** Define row labels ***
putexcel A6:B6 = "Means", merge hcenter bold
putexcel A7:B7 = "(s.d.)", merge hcenter bold
putexcel A8 = "Regressors", hcenter bold
putexcel B9 = "Class size", hcenter bold
putexcel B11 = "Percent disadvantaged", hcenter bold
putexcel B13 = "Enrollment", hcenter bold
putexcel B15 = "Enrollment squared/100", hcenter bold
putexcel B17 = "Piecewise linear trend", hcenter bold
putexcel B19 = "Root MSE", hcenter bold
putexcel B20 = "N", hcenter bold


*** Regression loop configuration ***
local samples "full_sample dis_sample"
local outcomes "average_verbal average_math"   // Three dependent variables
local regressors_1 "percent_disadvantaged"  
local regressors_2 "percent_disadvantaged enrollment"  
local regressors_3 "percent_disadvantaged enrollment enrollement_squared"  
local regressors_4 "trend"  
local columns_rf "C D E F"
local columns_rd "G H"
local columns_mf "I J K L"
local columns_md "M N"

save "${Data}\grades5_table4.dta", replace


*** Loop over grades ***
foreach outcome of local outcomes {

    foreach sampl of local samples {
		
		if ("`sampl'" == "full_sample") {
			use "${Data}\grades5_table4.dta", clear
			local cols = cond("`outcome'" == "average_verbal" & "`sampl'" == "full_sample", "`columns_rf'", "`columns_mf'")
			forval model = 1/4 {
				local col : word `model' of `cols' 				
					
					*** Select explanatory variables based on the model ***
				if `model' == 1 {
					local reg_controls "`regressors_1'"
				}
				else if `model' == 2 {
					local reg_controls "`regressors_2'"
				}
				else if `model' == 3 {
					local reg_controls "`regressors_3'"
				}
				else {
					local reg_controls "`regressors_4'"
				}

					*** Compute summary statistics ***
				summ `outcome', detail
				sca x = round(r(mean), 0.1)
				putexcel `col'6 = x, hcenter
				sca x = round(r(sd), 0.1)
				putexcel `col'7 = x, hcenter

				*** Run the regression ***
				ivregress 2sls `outcome' `reg_controls' (class_size=pred_size), cluster(schlcode)
				
				local reg_controls "class_size `reg_controls'" 
				local myvar `reg_controls'
				
				*** Store regression results ***
				local start_row = 9
				foreach var of local myvar {
					if ("`var'"=="trend") {
						sca x = round(_b[`var'], 0.001)
						putexcel `col'17 = x, hcenter
						local j = round(_se[`var'], 0.001)
						putexcel `col'18 = "( `j' )", hcenter
					}
					else {
						sca x = round(_b[`var'], 0.001)
						putexcel `col'`start_row' = x, hcenter
						local j = round(_se[`var'], 0.001)
						putexcel `col'`=`start_row' + 1' = "( `j' )", hcenter
						local start_row = `start_row' + 2
					}
				}

					*** Store model statistics ***
				sca x = round(e(rmse), 0.01)
				putexcel `col'19 = x, hcenter
				sca x = e(N)
				putexcel `col'20 = x, hcenter
			}
		}
		if ("`sampl'" == "dis_sample") {
			use "${Data}\grades5_table4.dta", clear
			keep if Enrollment==1
			local cols = cond("`outcome'" == "average_verbal" & "`sampl'" == "dis_sample", "`columns_rd'", "`columns_md'")
			forval model = 1/2 {
				local col : word `model' of `cols' 				

				*** Select explanatory variables based on the model ***
				if `model' == 1 {
					local reg_controls "`regressors_1'"
				}
				else {
					local reg_controls "`regressors_2'"
				}
					*** Compute summary statistics ***
				summ `outcome', detail
				sca x = round(r(mean), 0.1)
				putexcel `col'6 = x, hcenter
				sca x = round(r(sd), 0.1)
				putexcel `col'7 = x, hcenter

					*** Run the regression ***
				ivregress 2sls `outcome' `reg_controls' (class_size=pred_size), cluster(schlcode)

				local reg_controls "class_size `reg_controls'" 
				local myvar `reg_controls'
				
					*** Store regression results ***
				local start_row = 9
				foreach var of local myvar {
					if ("`var'"=="trend") {
						sca x = round(_b[`var'], 0.001)
						putexcel `col'17 = x, hcenter
						local j = round(_se[`var'], 0.001)
						putexcel `col'18 = "( `j' )", hcenter
					}
					else {
						sca x = round(_b[`var'], 0.001)
						putexcel `col'`start_row' = x, hcenter
						local j = round(_se[`var'], 0.001)
						putexcel `col'`=`start_row' + 1' = "( `j' )", hcenter
						local start_row = `start_row' + 2
					}
				}

				*** Store model statistics ***
				sca x = round(e(rmse), 0.01)
				putexcel `col'19 = x, hcenter
				sca x = e(N)
				putexcel `col'20 = x, hcenter
				
			}
        }
	}
}

putexcel set "$Outputs\Table4.xlsx", modify 

* Full sample
putexcel C6:F6, merge hcenter 
putexcel C7:F7, merge hcenter 
putexcel I6:L6, merge hcenter 
putexcel I7:L7, merge hcenter 
putexcel C20:E20, merge hcenter 
putexcel I20:K20, merge hcenter 


* Discontinuty sample
putexcel G6:H6, merge hcenter 
putexcel G7:H7, merge hcenter 
putexcel M6:N6, merge hcenter 
putexcel M7:N7, merge hcenter 
putexcel G20:H20, merge hcenter 
putexcel M20:N20, merge hcenter 


*** Final message ***
di "Table 4 successfully exported to Excel!"
restore