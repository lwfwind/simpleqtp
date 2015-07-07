An object oriented automation framework based on QTP. It is intended to make it easier to create a common platform to the automation engineer doing their job.

SimpleQTP can help maximize reusability of automation components and to minimize the maintenance costs to increase the efficiency, productivity, and quality of your automation testing.

Here are the features contained in this automation framework.

Provides a flow driver engine to make the scripts associated with test cases and test data, and determine which application/test suite/test case to run by Run Status flag.
Each test case execution script is designed as GUI Layer (the extension of an OR with dictionary object) and Business Layer (Actions), which provides a highly effective way that enables to exit a test smoothly, preventing QTP from getting stuck when it fails to identify GUI objects during runtime.
Provide a custom HTML report generator is designed based on XML+XSLT 1.0+CSS+Javascript technology which provides summary statistics and detail information(consists of clearly hierarchical tree ¨C Test Suite layer, Test Case layer and Test Step layer, and each layer can be expanded and collapsed).
Provide a rich, useful and generic function libraries (classes) to handle with such as database, excel, xml, FTP, registry, datetime, string.

To understand the current demo framework, please first read document files under Documentation
To run the demo, please change to correct URL/Location for flight4a.exe in PreExecutionSetup.xls file under Config folder, and then just double-click AutoRun.vbs.