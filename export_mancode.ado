*******************************************************************************
*  Version 0.3: 2023-05-02
*******************************************************************************
*  					Hannah Lachenmaier  &   Dennis Föste-Eggers
*
*  German Centre for Higher Education Research and Science Studies (DZHW)
*  Lange Laube 12, 30159 Hannover         
*  Phone: +49-(0)511 450 670-153		|	+49-(0)511 450 670-114
*  E-Mail (1): lachenmaier@dzhw.eu  	| 	foeste-eggers@dzhw.eu 
*  E-Mail (2):                          |   dennis.foeste@gmx.de
*
*******************************************************************************
*  Program name: export_mancode.ado     
*  Program purpose: Generates excel files containing respondents' answers 
*                   to be coded manually.
*******************************************************************************
*  Changes made:
*  Version 0.1: added GPL 
*  Version 0.2: added further optional macros  				       [2023-05-02]
*  Version 0.3: fixed minor bug regarding labels  				   [2023-05-03]
*******************************************************************************
*  License: GPL (>= 3)
*     
*	export_mancode.ado for Stata
*   Copyright (C) 2023 Lachenmaier, Hannah & Foeste-Eggers, Dennis
*
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*   GNU General Public License for more details.
*
*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*
*******************************************************************************
*  Citation: This code is © H. Lachenmaier & D. Foeste-Eggers, 2023, and it is  
*				made available under the GPL license enclosed with the software.
*
*!          Over and above the legal restrictions imposed by this license, if !
*!          you use this program for any (academic)	publication then you are  !
*!          obliged to provide proper attribution.                            !
*
*   H. Lachenmaier & D. Foeste-Eggers. export_mancode.ado for Stata, 
*           v0.3 (2023). [weblink].
*
*******************************************************************************

cap program drop export_mancode 
program define export_mancode  , nclass
    version 15
	
    syntax varlist [if] [in] using [, 			///
                                suffix_harm(namelist max=1) /// undocumented -tbd
                                suffix_source(namelist min=1 max=1) /// undocumented -tbd
                                suffix_code(namelist min=1 max=1) /// undocumented -tbd
                                keep_othvar(varlist) /// undocumented -tbd
                                suffix_sicherheit(namelist max=1) /// undocumented -tbd
                                suffix_geprueft(namelist max=1) /// undocumented -tbd
                                coder_var(namelist max=1) /// undocumented -tbd
                                comment_var(namelist max=1) /// undocumented -tbd
                                REPLACE 	            /// replace excel file
                                ]

                                
                                           

								
    di as txt "Excel file export via export_mancode.ado by Lachenmaier & Foeste-Eggers (2023,V0.3)"
    di ""
        
    
    *welche Variablen sollen exportiert werden
    local vlist = `"`varlist'"' // nicht mehr als 9 Variablen
    
    
    * welche Variablen sind noch fuer die Kodierung relevant und zu nutzen?
    local keep_othvar = `"`keep_othvar'"'
    
    * wo soll die Datei abgelegt werden
    local xlsx_export = `"`using'"'
    
    if `"`suffix_sicherheit'"'  == `""' local suffix_sicherheit = `"_sicherheit"'
    if `"`suffix_geprueft'"'    == `""' local suffix_geprueft = `"_geprueft"'
    if `"`coder_var'"'          == `""' local coder_var = `"codiert_von"'
    if `"`comment_var'"'        == `""' local comment_var = `"Bemerkungen"'
    
    if `"`suffix_harm'"' == `""' local suffix_harm = `"_harm"'
    if `"`suffix_code'"' == `""' local suffix_code = `"_code"'
    if `"`suffix_source'"' == `""' local suffix_source = `"_source"'
    
    ********************************************************************************
    ***	Behalte Beobachtungen, die nicht anhand der Sytax vercodet werden konnten **
    ********************************************************************************
    local count = 0
    foreach var of varlist `vlist' {
        local ++count
        if `count' ==1 {
            local vlist_cs = `"`var'`suffix_source'"'
            local vlist_s = `"`var'`suffix_source'"'
            local vlist_h = `"`var'`suffix_harm'"'
            * local vlist_a = `"`var'`asterisk'"'
            local vlist_a = `"`var'`suffix_harm' "' + ///
                            `"`var'`suffix_code' "' + ///
                            `"`var'`suffix_sicherheit' "' + ///
                            `"`var'`suffix_geprueft' "'
        }
        else {
            local vlist_cs = `"`vlist_cs', `var'`suffix_source'"'
            local vlist_s  = `"`vlist_s' `var'`suffix_source'"'
            local vlist_h  = `"`vlist_h' `var'_harm"'
            *local vlist_a  = `"`vlist_a' `var'`asterisk'"'
            local vlist_a = `"`vlist_a' `var'`suffix_harm' "' + ///
                            `"`var'`suffix_code' "' + ///
                            `"`var'`suffix_sicherheit' "' + ///
                            `"`var'`suffix_geprueft' "'
        }
    }
    preserve 
        qui { 
        if `"`if'`in'"' ~= `""' keep `if' `in' // hier möglich so aber eigentlich nicht    
        * noi di as result `"Temporary deletion of cases that are fully covered by automated coding"'
        keep if inlist(0,`vlist_cs')
        
        ********************************************************************************
        ***	Identifiziere moegliche Duplikate in strings und ggf. loesche diese	********
        *********************************************************************************
        duplicates list `vlist_h' `keep_othvar'
        
        duplicates drop `vlist_h' `keep_othvar', force
        *qui {
                
            
            ********************************************************************************
            ***	Generiere "leere" Variablen für manuelle Vercodung	************************
            *** Wie sicher mit Code-Vergabe?						************************
            ********************************************************************************
            *local suffix_sicherheit = `"_sicherheit"'
            
            foreach var of varlist `vlist' {
                gen `var'`suffix_sicherheit'= "" 
            *}
                sum `var'`suffix_source', meanonly 
                if `r(max)'>0 {
                    forvalues i = 0(1)`r(max)' {
                        local lab: label (`var'`suffix_source') `i'
                        replace `var'`suffix_sicherheit' = `"[`lab']"' if `var'`suffix_source' == `i'
                    }
                }
                
            *local suffix_geprueft = `"_geprueft"'
            
            *foreach var of varlist `vlist_h' {
                gen `var'`suffix_geprueft'=. 
            }
            
            *** Generiere Spalte: Wer hat codiert?
            gen `coder_var'  = .
            *** Generiere Spalte: Gibts Bemerkungen 
            gen `comment_var'= .
            
            ********************************************************************************
            *** Wandle Code-Variable in string, damit unterschiedliche missing-codes	****
            *** uebernommen werden (missing(repval) erscheint hier nicht sinnvoll)		****
            *** Falls system missing: replace zu leere Zelle							****
            ********************************************************************************
            foreach var of varlist `vlist' {
                tostring `var'`suffix_code', replace
                replace `var'`suffix_code'= "" if `var'`suffix_code'=="."
            }
            
            ********************************************************************************
            *** Behalte nur relevante Variablen fuer manuelle Vercodung	********************
            ********************************************************************************
            keep  `keep_othvar' `vlist_a' `coder_var' `comment_var'
            order `keep_othvar' `vlist_a' `coder_var' `comment_var'
            * drop `vlist'
            
            ********************************************************************************
            ***	Exportiere zu excel										********************
            ********************************************************************************
            * if `"`replace'"'~= `""' local replace = `", replace"'
            export excel `xlsx_export',  `replace' firstrow(variables)
            
            ********************************************************************************
            ***	Basic Formatierung der Excel-Liste						********************
            ********************************************************************************
            local xlsx_export = subinstr(`"`xlsx_export'"',`"using "',"",1)
            *noi di `"`xlsx_export'"'
            putexcel set `xlsx_export', modify
            putexcel A1:Z1, bold border(bottom, double)
            putexcel save
        }
    restore

    local xlsx_export = `xlsx_export'
    * di as txt ""
    di as smcl `"The following {stata `"!"`xlsx_export'""':xlsx file} has been generated"'
 
 end
    
