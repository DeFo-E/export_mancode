{smcl}
**!ado file version 0.3: 2023-05-04
help for {hi:export_mancode}{right: {browse "mailto:dennis.foeste@outlook.de":Dennis Foeste-Eggers}}
{hline}

{title:Title}
{phang}
{{bf:export_mancode} {hline 1} Generates excel file containing responses to be coded manually}

{title:Syntax}

{p 8 17 2}
{cmd:export_mancode} {varlist} {ifin} using [{cmd:,} {opth suffix_harm(namelist)} {opth suffix_source(namelist)}
                 {opth suffix_code(namelist)} {opth keep_othvar(varlist)} {opth suffix_sicherheit(namelist)}
                 {opth suffix_geprueft(namelist)} {opth coder_var(namelist)} {opth comment_var(namelist)} {opth replace}]
{p_end}


{title:Description}

{p}
{cmd:export_mancode} generates an excel file which contains harmonised responses that have to be coded manually and drops responses which are missing or have already been coded automatically. In order to make the coding process transparent, this programme adds further columns to enter the origin (source) of the codes, the certainty with which the code is assigned and the name of the person who assignes the codes. Additionally, the programme drops duplicates based on the harmonised strings. 
{p_end}

{title:Options}

{p 0 4}{cmd:suffix_harm} Include a column for each variable that is coded which captures the harmonised string of the response.
{p_end}

{p 0 4}{cmd:suffix_source} Include a column for each variable that is coded in which the source of the code is captured.
{p_end}

{p 0 4}{cmd:suffix_code} Inlcude a column for each variable in which the code of the string response is captured.
{p_end}

{p 0 4}{cmd:keep_othvar} Keep a set of additional pre-definded variables which contain context information for the further manual coding process. 
{p_end}

{p 0 4}{cmd:suffix_sicherheit} Include a column for each variable that is coded in which the certainty of the code assignment can be captured. 
{p_end}

{p 0 4}{cmd:suffix_geprueft} Include a column in which you can indicate whether the assigned code has been controlled by a third party. 
{p_end}


{p 0 4}{cmd:coder_var} Include a column in which the person who assignes codes can place his/her name. 
{p_end}

{p 0 4}{cmd:comment_var} Include a column in which the person who assignes codes can add further comments relevant for the coding process. 
{p_end}


{p 0 4}{cmd:replace} Replace the excel file. 
{p_end}