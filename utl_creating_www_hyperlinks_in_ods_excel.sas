Creating www hyperlinks in ods excel

  TWO SOLUTIONS

      1. ODS excel with active hyperlinks.
      2. ODS excel You have to activate hyperlinks

see SAS Forum
https://tinyurl.com/y9vmrk8m
https://communities.sas.com/t5/ODS-and-Base-Reporting/ODS-EXCEL-Clickable-external-hyperlink/m-p/494468?nobounce__

Surya Kiran  Profile
https://communities.sas.com/t5/user/viewprofilepage/user-id/83078

INPUT
=====

1. ODS excel with active hyperlinks.

 WORK.HAV1ST total obs=1
 ---------------------

     URLWWW            URLLNK

     Google    https://www.google.com/


2. ODS excel You have to activate hyperlinks
--------------------------------------------

 WORK.HAV2ND total obs=1
 -----------------------

    LINK                    DISPLAY_LINK

   Google    =HYPERLINK("https://google.com/","Google")

PROCESS
=======

 1. ODS excel with active hyperlinks.
 -------------------------------------

  ods excel file="d:/xls/utl_creating_www_hyperlinks_in_ods_excel_act.xlsx"
     options(sheet_name="google") style=statistical;
  proc report data=hav1st nowd ;
    columns urlLnk urlwww ;
    define urlLnk / display noprint;
    define urlwww / display "Link"
                  style={color=blue font_weight=bold font_size=12pt}
                 style(header)={color=black};
    compute urlwww;
       call define ('urlwww','URL',urlLnk);
    endcomp;

 run;quit;
 ods excel close;


2. ODS excel You have to activate hyperlinks
--------------------------------------------

  ods excel file="d:/xls/utl_creating_www_hyperlinks_in_ods_excel_noact.xlsx"
    options(sheet_name="google") style=statistical;

  proc print data=hav2nd noobs;
  run;quit;

  ods excel close;


OUTPUT
======

1. ODS excel with active hyperlinks.
------------------------------------

  SHEET CLASS IN WORKBOOK D:/XLS/CLASS.XLSX

  +--------------------+
  |          A         |
  +--------------------+
1 |        Link        |
  +--------------------+
2 |       GOOGLE       | * dark blue color
  +--------------------+

[GOOGLE]


2. ODS excel You have to activate hyperlinks
--------------------------------------------

  SHEET CLASS IN WORKBOOK D:/XLS/CLASS.XLSX

  You need to activate
  +--------------------+----------------------------------------------------+
  |          A         |                                                    |
  +--------------------+----------------------------------------------------+
1 |        Link        |          DISPLAY_LINK                              |
  +--------------------+----------------------------------------------------+
2 |       GOOGLE       |   =HYPERLINK("https://google.com/","Google")       |
  +--------------------+----------------------------------------------------+

[GOOGLE]

TWO METHODS TO ACTIVATE LINKS ABOVE

    1. Go to find replace and replace a blak with a blank
    2. Run the macro below (you need to leave the '=' off input


    /* T000227 ACTIVATING OLD STYLE HYPERLINKS IN EXCEL  */
    When you create hyperlinks from SAS leave off the first '=' sign, run this after to activate link
    Sub Hyp()
    Dim LastPlace, Z As Variant, X As Variant
    LastPlace = ActiveCell.SpecialCells(xlLastCell).Address
    ActiveSheet.Range(Cells(1, 1), LastPlace).Select
    Z = Selection.Address   'Get the address
        For Each X In ActiveSheet.Range(Z)  'Do while
            If Len(X) > 0 Then  'Find cells with something
                If Mid$(X.Text, 1, 5) = "HYPER" Then
                  X.FormulaR1C1 = "=" & Mid$(X.Text, 1)  '39 is code for tick
                End If
            Else
                X.FormulaR1C1 = ""  'If empty do not put tick
            End If
        Next
    End Sub

*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;


1. ODS excel with active hyperlinks.

   options nodate nonumber;title;footnote;
   data hav1st;
     urlwww='Google'; urlLnk='https://www.google.com/'; output;
   run;



2. ODS excel You have to activate hyperlinks
---------------------------------------------

  data hav2nd;
    Link="Google";
    Display_Link='=HYPERLINK("https://google.com/","Google")';
  run;quit;



