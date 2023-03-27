Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username, preperation

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKADM186-001 - Setup Warga Negara - View Detil Data.xlsx", "DPLKADM186-001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
preperation = Split(DataTable.Value("PREPERATION",dtlocalsheet),",")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, preperation)

Iteration = Environment.Value("ActionIteration")


REM ------- DPLK
Call DA_Login()
Browser("DPLK").Navigate "http://192.168.168.107/Account/Login?ReturnUrl=%2F" @@ hightlight id_;_788660_;_script infofile_;_ZIP::ssf8.xml_;_
Browser("DPLK").Page("Dashboard_2").Link("Setup").Click @@ script infofile_;_ZIP::ssf14.xml_;_
Browser("DPLK").Page("Dashboard_2").Link("Setup Umum").Click @@ script infofile_;_ZIP::ssf15.xml_;_
Browser("DPLK").Page("Dashboard_2").Link("Setup Warga Negara").Click @@ script infofile_;_ZIP::ssf16.xml_;_
Browser("DPLK").Page("Setup Warga Negara_2").WebEdit("WebEdit").Set "Vreemdelin" @@ script infofile_;_ZIP::ssf17.xml_;_
Browser("DPLK").Page("Setup Warga Negara_2").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("DPLK").Refresh @@ hightlight id_;_788660_;_script infofile_;_ZIP::ssf9.xml_;_
Browser("DPLK").Page("Setup Warga Negara - VIEW").Image("avatar1").Click @@ script infofile_;_ZIP::ssf19.xml_;_
Browser("DPLK").Page("Setup Warga Negara - VIEW").WebButton("Logout").Click @@ script infofile_;_ZIP::ssf20.xml_;_
Browser("DPLK").Page("Login_2").WebEdit("UserName").Set "32074" @@ script infofile_;_ZIP::ssf10.xml_;_
Browser("DPLK").Page("Login_2").WebEdit("Password").SetSecure "63e060f8746033a45b0f8b95d5e31d1604cecf26" @@ script infofile_;_ZIP::ssf11.xml_;_
Browser("DPLK").Page("Login_2").WebButton("Login").Click @@ script infofile_;_ZIP::ssf12.xml_;_
Browser("DPLK").Page("Dashboard_2").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf13.xml_;_
Call AC_GoTo_Menu()


	Call Lihat_Setup_Administration_Setup_Warga_Negara()


Call DA_Logout("0")
Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDPLKPath 	= Environment.Value("TestDir")
	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	
	LibPathDPLK	= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport			= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDPLK & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Administration_Setup.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_LogMenu.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Dashboard.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Log.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Administration_Setup.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Username					= DataTable.Value("USERID",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
End Sub
