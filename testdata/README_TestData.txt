HMA - test data (anonymised, no information about a real excavation)
====================================================================

HMA_TestData_EN.csv contains 95 stratigraphic units with 122 "lies above / lies below" relations and
5 contemporary pairs. All IDs, names, periods, phases and texts are neutral:

  - IDs run from XTESTX_1 to XTESTX_95 (a few numbers are unused), plus two special units:
    XTESTX_0 is the topmost unit (no other unit lies above it) and
    XTESTX_9999 is the lowest unit (it lies on no other unit).
  - 5 periods, 10 phases, 5 categories (Deposit, Fill, Levelling, Wall, Interface),
    18 units without a period and 32 units with finds.
  - The file is a text table (CSV) with the delimiter ";" and English column names.

Try it
------
  1. Start HMA. Choose "File (table / database)" as the source, click "Choose file..." and select
     HMA_TestData_EN.csv.
  2. Tab "Columns": click "Detect columns automatically". All nine columns are mapped correctly.
  3. Tab "Values": tick "Interface" in the list of category values (these units are interfaces;
     all other categories are deposits).
  4. Look at the tabs "Table" and "Preview", then click "Create matrix".

Optional: add the nodes T and G
-------------------------------
XTESTX_0 and XTESTX_9999 are meant for the optional nodes "Top surface" (T) and "Interface to Geology" (G):

  5. Tab "Values", section "Optional: insert T, G and U":
       - tick "Insert Top surface (T)" and select XTESTX_0 in the list below it;
       - tick "Insert Interface to Geology (G)" and select XTESTX_9999 in the list below it.
     In the "Preview", T now sits above XTESTX_0 and G below XTESTX_9999.
     T and G exist only in the HMC output; HME has no such nodes, so HMA leaves them out there.

Expected result
---------------
  HMC:  95 nodes, 5 parent groups, 10 sub-groups, 127 edges
        (with T and G: 97 nodes, 129 edges)
  HME:  122 edges (HME has no "contemporary" relations)
The log reports no errors.
