# utl-altair-slc-adding-a-sheet-to-a-xlsm-workbook
Altair slc adding a sheet to a xlsm workbook
    %let pgm=utl-altair-slc-adding-a-sheet-to-a-xlsm-workbook;

    %stop_submission;

    The posted response stated that Monarch does not support exports to a macro-enabled workbook;

    RE: Altair slc adding a sheet to a xlsm workbook

    Too long to post in listserv, see github

    gihub
    https://github.com/rogerjdeangelis/utl-altair-slc-adding-a-sheet-to-a-xlsm-workbook

    community.altair.com
    https://community.altair.com/discussion/6403

     SOLUTION

        1 CREATE XLSM WORKBOOK WITH MACRO. SUM A2-A21

          Sub sum_weight()
              Range("A21").Select
              ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
              Range("A22").Select
          End Sub


        2 ADD A SECOND WORKSHEET CLASS TO XLSM WORKBOOK


     INPUT d:/xls/d:/xls/workbook_with_macro.xlsm

    /********************************************************************************************/
    /* PROBLEM: ADD SHEET CLASS                |                                                */
    /*                                         |                                                */
    /*             INPUT                       |            OUTPUT (added sheet)                */
    /*                                         |                                                */
    /* d:/xls/d:/xls/workbook_with_macro.xlsm  |                                                */
    /* If you run the macro it will sum A2-A20 | d:/xls/d:/xls/workbook_with_macro.xlsm add     */
    /*                                         | sdded sheet CLASS                              */
    /*      ----------------------+            | -----------------------------+                 */
    /*      |A1| fx         |     |            | |A1| fx     | NAME           |                 */
    /*      -----------------------            | --------------------------------------------+  */
    /*      [ ]|     A      |     |            | [ ]|     A   |  B  |   C  |   D    |   E    |  */
    /*      -  +-------------------            | -  +----------------------------------------+  */
    /*      1  |            |     |            | 1  | NAME    | SEX |  AGE | HEIGHT | WEIGHT |  */
    /*      -  +------------+-----+            | -  +---------+-----+------+--------+--------+  */
    /*      2  |     1      |     |            | 2  | ALFRED  |  M  |  99  |   69   | 112.5  |  */
    /*      -  +------------+-----+            | -  +---------+-----+------+--------+--------+  */
    /*      3  |     2      |     |            | 3  | BARBARA |  F  |  13  |   58   | 101.5  |  */
    /*      -  +------------+-----+            | -  +---------+-----+------+--------+--------+  */
    /*          ...                            |     ...                                        */
    /*         +------------+-----+            |    +---------+-----+------+--------+--------+  */
    /*      20 |    200     |     |            | 20 | WILLIAM |  M  |  15  |  66.5  | 112    |  */
    /*         +------------+-----+            |    +---------+-----+------------------------+  */
    /*      21 |            |     |            | [CLASS]                                        */
    /*         +------------+-----+            |                                                */
    /*      [sheet]                            |                                                */
    /********************************************************************************************/

    /*                      _              _                 _                   _
    / |  ___ _ __ ___  __ _| |_ ___  __  _| |___ _ __ ___   (_)_ __  _ __  _   _| |_
    | | / __| `__/ _ \/ _` | __/ _ \ \ \/ / / __| `_ ` _ \  | | `_ \| `_ \| | | | __|
    | || (__| | |  __/ (_| | ||  __/  >  <| \__ \ | | | | | | | | | | |_) | |_| | |_
    |_| \___|_|  \___|\__,_|\__\___| /_/\_\_|___/_| |_| |_| |_|_| |_| .__/ \__,_|\__|
                                                                    |_|
    */

    %utlfkil(d:/xls/tempx_workbook.xlsx);
    %utlfkil(d:/xls/workbook_with_macro.xlsm);

    options set=PYTHONHOME "D:\python310";
    proc python;
    submit;

    import openpyxl
    import xlwings as xw
    import os

    def create_xlsm_with_macro():
        # Create a new workbook
        wb = openpyxl.Workbook()
        ws = wb.active

        # Add some sample data in column A (rows 2-20)
        for i in range(2, 21):
            ws[f'A{i}'] = i * 10  # Sample data

        # Save as regular xlsx first
        temp_file = 'd:/xls/tempx_workbook.xlsx'
        wb.save(temp_file)
        wb.close()

        # Use xlwings to open and add macro
        app = xw.App(visible=False)
        wb_xw = app.books.open(temp_file)

        # Add VBA module with macro
        vba_code = '''
    Sub sum_weight()
        Range("A21").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
        Range("A22").Select
    End Sub
    '''

        # Create module and add code
        macro_module = wb_xw.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        macro_module.CodeModule.AddFromString(vba_code)

        # Save as xlsm
        wb_xw.save('d:/xls/workbook_with_macro.xlsm')
        wb_xw.close()
        app.quit()

        # Clean up temporary file
        os.remove(temp_file)
        print("Workbook with macro created successfully!")

    create_xlsm_with_macro()
    endsubmit;
    run;quit;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    d:/xls/d:/xls/workbook_with_macro.xlsm
    If you run the macro it will sum A2-A20

         ----------------------+
         |A1| fx         |     |
         -----------------------
         [ ]|     A      |     |
         -  +-------------------
         1  |            |     |
         -  +------------+-----+
         2  |     1      |     |
         -  +------------+-----+
         3  |     2      |     |
         -  +------------+-----+
             ...
            +------------+-----+
         20 |    200     |     |
            +------------+-----+
         21 |            |     |
            +------------+-----+
         [sheet]

    And Macro

    Sub sum_weight()
        Range("A21").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
        Range("A22").Select
    End Sub


    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       15:29 Monday, November 10, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"

    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.024
          cpu time  : 0.015


    NOTE: AUTOEXEC processing completed

    1
    2          %utlfkil(d:/xls/tempx_workbook.xlsx);
    The file d:/xls/tempx_workbook.xlsx does not exist
    3         %utlfkil(d:/xls/workbook_with_macro.xlsm);
    The file d:/xls/workbook_with_macro.xlsm does not exist
    4
    5         options set=PYTHONHOME "D:\python310";
    6         proc python;
    7         submit;
    8
    9         import openpyxl
    10        import xlwings as xw
    11        import os
    12
    13        def create_xlsm_with_macro():
    14            # Create a new workbook
    15            wb = openpyxl.Workbook()
    16            ws = wb.active
    17
    18            # Add some sample data in column A (rows 2-20)
    19            for i in range(2, 21):
    20                ws[f'A{i}'] = i * 10  # Sample data
    21
    22            # Save as regular xlsx first
    23            temp_file = 'd:/xls/tempx_workbook.xlsx'
    24            wb.save(temp_file)
    25            wb.close()
    26
    27            # Use xlwings to open and add macro
    28            app = xw.App(visible=False)
    29            wb_xw = app.books.open(temp_file)
    30
    31            # Add VBA module with macro

    2                                          Altair SLC       15:29 Monday, November 10, 2025

    32            vba_code = '''
    33        Sub sum_weight()
    34            Range("A21").Select
    35            ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
    36            Range("A22").Select
    37        End Sub
    38        '''
    39
    40            # Create module and add code
    41            macro_module = wb_xw.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    42            macro_module.CodeModule.AddFromString(vba_code)
    43
    44            # Save as xlsm
    45            wb_xw.save('d:/xls/workbook_with_macro.xlsm')
    46            wb_xw.close()
    47            app.quit()
    48
    49            # Clean up temporary file
    50            os.remove(temp_file)
    51            print("Workbook with macro created successfully!")
    52
    53        create_xlsm_with_macro()
    54        endsubmit;

    NOTE: Submitting statements to Python:


    55        run;quit;
    NOTE: Procedure python step took :
          real time : 3.264
          cpu time  : 0.015


    56
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 3.362
          cpu time  : 0.062

    /*___              _     _                      _        _               _         _
    |___ \    __ _  __| | __| | __      _____  _ __| | _____| |__   ___  ___| |_   ___| | __ _ ___ ___
      __) |  / _` |/ _` |/ _` | \ \ /\ / / _ \| `__| |/ / __| `_ \ / _ \/ _ \ __| / __| |/ _` / __/ __|
     / __/  | (_| | (_| | (_| |  \ V  V / (_) | |  |   <\__ \ | | |  __/  __/ |_ | (__| | (_| \__ \__ \
    |_____|  \__,_|\__,_|\__,_|   \_/\_/ \___/|_|  |_|\_\___/_| |_|\___|\___|\__| \___|_|\__,_|___/___/
    */


    libname xls excel "d:/xls/workbook_with_macro.xlsm";

    data xls.class;
       informat
         NAME $8.
         SEX $1.
         AGE 8.
         HEIGHT 8.
         WEIGHT 8.
    ;
    input NAME SEX AGE HEIGHT WEIGHT;
    cards4;
    Alfred M 14 69 112.5
    Alice F 13 56.5 84
    Barbara F 13 65.3 98
    Carol F 14 62.8 102.5
    Henry M 14 63.5 102.5
    James M 12 57.3 83
    Jane F 12 59.8 84.5
    Janet F 15 62.5 112.5
    Jeffrey M 13 62.5 84
    John M 12 59 99.5
    Joyce F 11 51.3 50.5
    Judy F 14 64.3 90
    Louise F 12 56.3 77
    Mary F 15 66.5 112
    Philip M 16 72 150
    Robert M 12 64.8 128
    Ronald M 15 67 133
    Thomas M 11 57.5 85
    William M 15 66.5 112
    ;;;;
    run;quit;

    libname xls clear;

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
