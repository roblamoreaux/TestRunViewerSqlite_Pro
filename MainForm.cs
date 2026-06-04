using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Microsoft.Data.Sqlite;
using System.Data;
using System.Text;
using System.Text.Json;
//using SqliteFunctions;

namespace TestRunViewerSqlite
{
    public partial class MainForm : Form
    {
        private string _dbPath = string.Empty;
        private string ConnectionString => $"Data Source={_dbPath};Mode=ReadOnly;Cache=Shared"; private string SettingsFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TestRunViewerSqlite"); private string SettingsFile => Path.Combine(SettingsFolder, "settings.json");

        /*        private string ConnectionString => $"Data Source={_dbPath};Mode=ReadOnly;Cache=Shared"; 
                private string SettingsFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TestRunViewerSqlite"); 
                private string SettingsFile => Path.Combine(SettingsFolder, "settings.json");
        */
        private readonly string _detailsSql = @"-- GetParentName source

--WITH 
WITH RECURSIVE GetStepName as (
SELECT
	testrun.RunID , 
	TestRun.ParentRunID, 
	testRun.PlanRunID ,
	params.Value As ParentName,
	'' as StepName,
	0 as level
FROM
	(testrun
INNER JOIN testrun2params ON
	((testrun.RunID = testrun2params.Scope)
		AND (testrun.runid = testrun2params.runid))
INNER JOIN params ON
		testrun2params.paramid = params.paramid)
WHERE
	((((params.Name)= 'Name')
		AND ((params.GroupName)= '')))
	AND testrun.ParentRunID is NULL
UNION
SELECT
	e.RunID ,
	e.ParentRunID,
	e.PlanRunID,
	eh.ParentName || '\' || eh.StepName ,
	params.Value ,
	eh.level + 1 as level
FROM
	testrun e
join GetStepName eh ON
	e.ParentRunID = eh.RunID
INNER JOIN testrun2params ON
	( (e.runid = testrun2params.runid )
		AND (e.runid = testrun2params.Scope))
INNER JOIN params ON
		testrun2params.paramid = params.paramid
WHERE
	((((params.Name)= 'Name')
		AND ((params.GroupName)= '')))
),
--with 
GetStepMinLimit AS
(
SELECT
	testrun2params.scope,
	testrun2params.runid,
	params.groupname,
	params.name,
	params.value AS MinLimit
FROM
	(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.groupname)= 'Limits')
				And ((params.name)= 'Minimum Value')))
,
 GetStepMaxLimit AS 
(
SELECT
	testrun2params.scope,
	testrun2params.runid,
	params.name,
	params.value AS MaxLimit,
	params.groupname
FROM
	(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.name)= 'Maximum Value')
				And ((params.groupname)= 'Limits')))
,
GetStepVerdict AS (
SELECT
	testrun2params.scope,
	testrun2params.runid,
	params.name,
	params.value AS Verdict
FROM
	(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.name)= 'Verdict')
				And ((params.value)<> 'NotSet')))
,
GetStepTimestamp AS (
SELECT
	testrun2params.scope,
	testrun2params.runid,
	params.name,
	params.value AS [TimeStamp],
	params.groupname
FROM
	(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.name)= 'StartTime')))
,
 GetStepCheckLimit AS(
SELECT testrun2params.scope, testrun2params.runid, params.name, params.value AS CheckLimit, params.groupname
FROM (testrun INNER JOIN testrun2params ON testrun.runid = testrun2params.runid) INNER JOIN params ON testrun2params.paramid = params.paramid
WHERE (((testrun2params.scope)= testrun.runid) And ((testrun2params.runid)= testrun.runid) And ((params.name)= 'Check Limits') And ((params.groupname)= 'Limits')))
--,
/* GetStepName AS(
SELECT testrun2params.scope, testrun2params.runid, params.name, params.value AS StepName, params.groupname
FROM (testrun INNER JOIN testrun2params ON testrun.runid = testrun2params.runid) INNER JOIN params ON testrun2params.paramid = params.paramid
WHERE (((testrun2params.scope)= testrun.runid) And ((testrun2params.runid)= testrun.runid) And ((params.name)= 'Step Name')))
,
 GetParentName AS(
SELECT testrun.runid, testrun.parentrunid, testrun2params.scope, params.Value AS ParentName
FROM (testrun INNER JOIN testrun2params ON (testrun.ParentRunID = testrun2params.Scope) AND (testrun.runid = testrun2params.runid)) INNER JOIN params ON testrun2params.paramid = params.paramid
WHERE (((params.Name)= 'Name') AND ((params.GroupName)= ''))
),
GrandParentName as (
SELECT
	testrun.runid as gprunid,
	testrun.parentrunid as gpprunid,
	testrun2params.scope as gpscope,
	params.Value AS GParentName
FROM
	(testrun
INNER JOIN testrun2params ON
	(testrun.ParentRunID = testrun2params.Scope)
	AND (testrun.runid = testrun2params.runid))
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((params.Name)= 'Name')
		AND ((params.GroupName)= ''))
),
GGrandParentName as (
SELECT
	testrun.runid as gprunid,
	testrun.parentrunid as gpprunid,
	testrun2params.scope as gpscope,
	params.Value AS GGParentName
FROM
	(testrun
INNER JOIN testrun2params ON
	(testrun.ParentRunID = testrun2params.Scope)
	AND (testrun.runid = testrun2params.runid))
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((params.Name)= 'Name')
		AND ((params.GroupName)= ''))
)*/
SELECT
	--GGrandParentName.GGParentName,
	--GrandParentName.GParentName,
	--GetParentName.ParentName,,
    --testrun.RunID || '|' || GetStepTimestamp.TimeStamp || '|' || COALESCE(result.Dim0, 'NULL') AS PrimaryKey,
	GetStepName.ParentName,
	GetStepName.StepName,
	GetStepTimestamp.TimeStamp,
	resulttype.Dim0 as SignalName,
	result.dim0 as Signal,
	resulttype.Dim1 as ValueName,
	result.dim1 as Value,
	GetStepMinLimit.MinLimit,
	GetStepMaxLimit.MaxLimit,
	GetStepVerdict.Verdict,
	GetStepCheckLimit.CheckLimit,
	resulttype.Dim2 as ValueName2,
	result.Dim2 as Value2,
	resulttype.Dim3 as ValueName3,
	result.Dim3 as Value3,
	planrun.planrunnumber,
	testrun.parentrunid,
	testrun.runid,
    testrun.RunID || '|' || GetStepTimestamp.TimeStamp || '|' || COALESCE(result.Dim0, 'NULL') AS PrimaryKey
	
	
FROM
	--((
	((((((((((planrun
left JOIN testrun ON
	planrun.runid = testrun.planrunid)
LEFT JOIN resultseries ON
	testrun.runid = resultseries.runid)
LEFT JOIN result ON
	resultseries.resultseriesid = result.resultseriesid)
LEFT JOIN resulttype ON
	resultseries.resulttypeid = resulttype.resulttypeid)
LEFT JOIN GetStepName ON
	testrun.runid = GetStepName.runid)
LEFT JOIN GetStepMinLimit ON
	testrun.runid = GetStepMinLimit.runid)
LEFT JOIN GetStepMaxLimit ON
	testrun.runid = GetStepMaxLimit.runid)
LEFT JOIN GetStepVerdict ON
	testrun.runid = GetStepVerdict.runid)
LEFT JOIN GetStepTimestamp ON
	testrun.runid = GetStepTimestamp.runid)
LEFT JOIN GetStepCheckLimit ON
	testrun.RunID = GetStepCheckLimit.runid)
/*LEFT JOIN GetParentName ON
	testrun.RunID = GetParentName.runid)
 Left Join GrandParentName on GetParentName.parentrunid = GrandParentName.gprunid)  
 Left Join GGrandParentName on GrandParentName.gpprunid  = GGrandParentName.gprunid
*/
	WHERE
	(((planrun.planrunnumber)=@PlanRunNumber))
---ORDER BY
GROUP BY 
    GetStepName.ParentName,
	GetStepName.StepName,
	GetStepTimestamp.TimeStamp,
	resulttype.Dim0,
	result.dim0,
	resulttype.Dim1,
	resulttype.Dim2,
	resulttype.Dim3,
	planrun.planrunnumber,
	testrun.parentrunid,
	testrun.runid,
    PrimaryKey
ORDER BY
	PrimaryKey;";
/*    testrun.runid,
	GetStepTimestamp.TimeStamp,
    resulttype.Dim0,
	result.dim0;";
*/
        private readonly string _overviewSqlSN = @"-- GetParentName source

WITH 
GetStepTimestamp AS (
SELECT
	testrun2params.scope,
	testrun2params.runid,
	params.name,
	params.value AS TimeStamp,
	params.groupname
FROM
	(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.name)= 'StartTime')))
,
 GetStepName AS
 (
SELECT
	testrun2params.scope,
	testrun2params.runid,
	params.name,
	params.value AS StepName,
	params.groupname
FROM
	(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.name)= 'Step Name')))
,
 GetVerdict AS
 (
SELECT
	planrun.planrunnumber,
	params.value as Verdict
FROM
	planrun
INNER JOIN ((testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid) ON
	planrun.runid = testrun.runid
WHERE
	(((testrun2params.scope)= testrun.runid)
		And ((testrun2params.runid)= testrun.runid)
			And ((params.name)= 'Verdict'))
),
 GetParentName AS
 (
SELECT
	testrun.runid,
	testrun.parentrunid,
	testrun2params.scope,
	params.Value AS ParentName
FROM
	(testrun
INNER JOIN testrun2params ON
	(testrun.ParentRunID = testrun2params.Scope)
		AND (testrun.runid = testrun2params.runid))
INNER JOIN params ON
	testrun2params.paramid = params.paramid
WHERE
	(((params.Name)= 'Name')
		AND ((params.GroupName)= '')))
SELECT
	GetStepTimestamp.TimeStamp,
	GetParentName.ParentName as PlanRunName,
	planrun.planrunnumber,
	result.dim0 as Signal,
	result.dim1 as SerialNumber,
	GetVerdict.Verdict,
	testrun.runid
FROM
	(((((((planrun
INNER JOIN testrun ON
	planrun.runid = testrun.planrunid)
LEFT JOIN GetStepTimestamp ON
	testrun.runid = GetStepTimestamp.runid)
LEFT JOIN GetParentName ON
	testrun.RunID = GetParentName.runid
LEFT JOIN GetVerdict ON
	planrun.planrunnumber = GetVerdict.planrunnumber)
LEFT JOIN resultseries ON                                                                                                                                                                                                                                        
	testrun.runid = resultseries.runid)
LEFT JOIN result ON
	resultseries.resultseriesid = result.resultseriesid)
LEFT JOIN resulttype ON
	resultseries.resulttypeid = resulttype.resulttypeid)
JOIN GetStepName ON
	testrun.runid = GetStepName.runid)
WHERE
	( --((GetStepName.StepName = 'Operator Inputs Receiver') or (GetStepName.StepName = 'User Input Test Info Step') or (GetStepName.StepName = 'Factory Data Test Info Step'))
		--and 
		((Signal = 'Serial Number') )
)
ORDER BY
	planrun.PlanRunNumber ,
	GetStepTimestamp.TimeStamp
";


        private readonly string _overviewSql = @"
WITH 
GetStepTimestamp AS (
  SELECT testrun2params.scope, testrun2params.runid, params.value AS TimeStamp
  FROM testrun
  INNER JOIN testrun2params ON testrun.runid = testrun2params.runid
  INNER JOIN params ON testrun2params.paramid = params.paramid
  WHERE testrun2params.scope = testrun.runid AND testrun2params.runid = testrun.runid AND params.name = 'StartTime'
),
GetParentName AS (
  SELECT testrun.runid, testrun.parentrunid, params.Value AS ParentName
  FROM testrun
  INNER JOIN testrun2params ON testrun.ParentRunID = testrun2params.Scope AND testrun.runid = testrun2params.runid
  INNER JOIN params ON testrun2params.paramid = params.paramid
  WHERE params.Name = 'Name' AND params.GroupName = ''
),
GetVerdict AS (
  SELECT planrun.planrunnumber, params.value as Verdict
  FROM planrun
  INNER JOIN testrun ON planrun.runid = testrun.runid
  INNER JOIN testrun2params ON testrun.runid = testrun2params.runid
  INNER JOIN params ON testrun2params.paramid = params.paramid
  WHERE testrun2params.scope = testrun.runid AND testrun2params.runid = testrun.runid AND params.name = 'Verdict'
)
SELECT GetStepTimestamp.TimeStamp, GetParentName.ParentName as PlanRunName, planrun.PlanRunNumber, GetVerdict.Verdict, planrun.runid
FROM planrun
INNER JOIN testrun ON planrun.runid = testrun.planrunid
LEFT  JOIN GetVerdict ON planrun.planrunnumber = GetVerdict.planrunnumber
LEFT  JOIN GetParentName ON testrun.RunID = GetParentName.runid
INNER JOIN GetStepTimestamp ON testrun.runid = GetStepTimestamp.runid
WHERE testrun.RunID = (PlanRun.RunID + 1)
ORDER BY planrun.PlanRunNumber, GetStepTimestamp.TimeStamp
";
        private readonly string _summarySql = @"SELECT
	planrun.planrunnumber,
	params.groupname,
	params.name,
	params.value,
    params.paramid
FROM
	planrun
INNER JOIN (
(testrun
INNER JOIN testrun2params ON
	testrun.runid = testrun2params.runid)
INNER JOIN params ON
	testrun2params.paramid = params.paramid) ON
	planrun.runid = testrun.runid
where
	(planrun.planrunnumber = @PlanRunNumber);"; 
        private readonly string _statsAllSql = @"
-- Get Stats source
With  GetReport AS (
		WITH RECURSIVE GetStepName as (
		SELECT
			testrun.RunID , 
			TestRun.ParentRunID, 
			testRun.PlanRunID ,
			params.Value As ParentName,
			'' as StepName,
			0 as level
		FROM
			(testrun
		INNER JOIN testrun2params ON
			((testrun.RunID = testrun2params.Scope)
				AND (testrun.runid = testrun2params.runid))
		INNER JOIN params ON
				testrun2params.paramid = params.paramid)
		WHERE
			((((params.Name)= 'Name')
				AND ((params.GroupName)= '')))
			AND testrun.ParentRunID is NULL
		UNION
		SELECT
			e.RunID ,
			e.ParentRunID,
			e.PlanRunID,
			eh.ParentName || '\' || eh.StepName ,
			params.Value ,
			eh.level + 1 as level
		FROM
			testrun e
		join GetStepName eh ON
			e.ParentRunID = eh.RunID
		INNER JOIN testrun2params ON
			( (e.runid = testrun2params.runid )
				AND (e.runid = testrun2params.Scope))
		INNER JOIN params ON
				testrun2params.paramid = params.paramid
		WHERE
			((((params.Name)= 'Name')
				AND ((params.GroupName)= '')))
		),
	  GetStepMinLimit AS
		(
		SELECT
			testrun2params.scope,
			testrun2params.runid,
			params.groupname,
			params.name,
			params.value as MinLimit
            --case
            --   WHEN
            --     ((typeof(params.value) = 'integer')
            --      OR  (typeof(params.value) = 'real'))
            --   THEN
            --     -999999.00 --params.value
            --   ELSE
            --     0.0 
            --END MinLimit
		FROM
			(testrun
		INNER JOIN testrun2params ON
			testrun.runid = testrun2params.runid)
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((testrun2params.scope)= testrun.runid)
				And ((testrun2params.runid)= testrun.runid)
					And ((params.groupname)= 'Limits')
						And ((params.name)= 'Minimum Value')))
		,
	 GetStepMaxLimit AS 
		(
		SELECT
			testrun2params.scope,
			testrun2params.runid,
			params.name,
			params.value  AS MaxLimit,
            params.groupname
		FROM
			(testrun
		INNER JOIN testrun2params ON
			testrun.runid = testrun2params.runid)
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((testrun2params.scope)= testrun.runid)
				And ((testrun2params.runid)= testrun.runid)
					And ((params.name)= 'Maximum Value')
						And ((params.groupname)= 'Limits')))
	,
	GetStepVerdict AS (
		SELECT
			testrun2params.scope,
			testrun2params.runid,
			params.name,
			params.value AS Verdict
		FROM
			(testrun
		INNER JOIN testrun2params ON
			testrun.runid = testrun2params.runid)
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((testrun2params.scope)= testrun.runid)
				And ((testrun2params.runid)= testrun.runid)
					And ((params.name)= 'Verdict')
						And ((params.value)<> 'NotSet')))
		,
		GetStepTimestamp AS (
		SELECT
			testrun2params.scope,
			testrun2params.runid,
			params.name,
			params.value AS [TimeStamp],
			params.groupname
		FROM
			(testrun
		INNER JOIN testrun2params ON
			testrun.runid = testrun2params.runid)
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((testrun2params.scope)= testrun.runid)
				And ((testrun2params.runid)= testrun.runid)
					And ((params.name)= 'StartTime')))
	,
	 GetStepCheckLimit AS (
		SELECT
			testrun2params.scope,
			testrun2params.runid,
			params.name,
			params.value AS CheckLimit,
			params.groupname
		FROM
			(testrun
		INNER JOIN testrun2params ON
			testrun.runid = testrun2params.runid)
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((testrun2params.scope)= testrun.runid)
				And ((testrun2params.runid)= testrun.runid)
					And ((params.name)= 'Check Limits')
						And ((params.groupname)= 'Limits'))
	)
	SELECT
		GetStepName.ParentName,
		GetStepName.StepName,
		GetStepTimestamp.TimeStamp,
		resulttype.Dim0 as SignalName,
		Case
			WHEN 
				((resulttype.Dim0 = 'Measurement')
					OR (resulttype.Dim0 = 'Calibration')
					Or (resulttype.Dim0 = 'Source')
					OR (resulttype.Dim0 = 'Signal')
					OR (resulttype.Dim0 = 'Value')
					)
				THEN
					COALESCE(result.Dim0, 'Result.dim0' )----cast (result.Dim0 as TEXT)
			WHEN
			   (resulttype.Dim0 = 'IPAddress')
				THEN 
				    'IPADR' ----cast ('IPADDRESS' as TEXT)
            ELSE
				COALESCE(resulttype.Dim0, 'NULL')
		END ValueName,
		result.dim0 as Signal,
		--resulttype.Dim1 as ValueName,
		--result.dim1 as Value,
		Case
			WHEN 
				((resulttype.Dim0 = 'Measurement')
					OR (resulttype.Dim0 = 'Calibration')
					Or (resulttype.Dim0 = 'Source')
					OR (resulttype.Dim0 = 'Signal')
					OR (resulttype.Dim0 = 'Value')
					)
				THEN
					cast (result.dim1  as REAL)
			WHEN
			   (resulttype.Dim0 = 'IPAddress')
				THEN 
				    cast (0.0   as REAL)
            WHEN ((resulttype.DIM0 = NULL)
                 OR (result.Dim0 = NULL )) 
                THEN 
                  cast (0.0   as REAL)
			ELSE
				COALESCE(result.Dim0, 0.0) --cast (result.dim0    as REAL)
		END Value,
		COALESCE(GetStepMinLimit.MinLimit, 0.0)  as MinLimit,
		COALESCE(GetStepMaxLimit.MaxLimit, 0.0)  as MaxLimit,
		GetStepVerdict.Verdict,
		GetStepCheckLimit.CheckLimit,
		resulttype.Dim2 as ValueName2,
		result.Dim2 as Value2,
		resulttype.Dim3 as ValueName3,
		result.Dim3 as Value3,
		planrun.planrunnumber,
		testrun.runid
	FROM
		((((((((((planrun
	INNER JOIN testrun ON
		planrun.runid = testrun.planrunid)
	LEFT JOIN resultseries ON
		testrun.runid = resultseries.runid)
	LEFT JOIN result ON
		resultseries.resultseriesid = result.resultseriesid)
	LEFT JOIN resulttype ON
		resultseries.resulttypeid = resulttype.resulttypeid)
	LEFT JOIN GetStepName ON
		testrun.runid = GetStepName.runid)
	LEFT JOIN GetStepMinLimit ON
		testrun.runid = GetStepMinLimit.runid)
	LEFT JOIN GetStepMaxLimit ON
		testrun.runid = GetStepMaxLimit.runid)
	LEFT JOIN GetStepVerdict ON
		testrun.runid = GetStepVerdict.runid)
	LEFT JOIN GetStepTimestamp ON
		testrun.runid = GetStepTimestamp.runid)
	LEFT JOIN GetStepCheckLimit ON
		testrun.RunID = GetStepCheckLimit.runid)
	),
	 GetScriptName AS
		 (
		SELECT
			testrun.planrunid,
			testrun.runid,
			testrun.parentrunid,
			testrun2params.scope,
			params.Value AS ScriptName
		FROM
			(testrun
		INNER JOIN testrun2params ON
			(testrun.PlanRunID = testrun2params.Scope)
				AND (testrun.runid = testrun2params.runid))
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((params.Name)= 'Name')
				AND ((params.GroupName)= ''))
		)
Select
	GetScriptName.ScriptName, 
	GetReport.ParentName, 
	GetReport.StepName,
	GetReport.ValueName,
	Verdict,
	COALESCE(avg(GetReport.Value),0.0) as Average,
	cast (COALESCE(Min(GetReport.Value),0.0) as real) as Minimum,
	cast (COALESCE(Max(GetReport.Value),0.0) as real) as Maximum,
	variance(GetReport.Value) as Variance,
	COALESCE(stdev(GetReport.Value),0.0) as StandardDeviation,
	Count(GetReport.Value) as Count,
	GetReport.MinLimit,
	GetReport.MaxLimit,
	--min((GetReport.MaxLimit - avg(GetReport.Value))/(3*stdev(GetReport.Value)), (avg(GetReport.Value)- GetReport.MinLimit)/(3*stdev(GetReport.Value))) as CPK,
	--((GetReport.MaxLimit - avg(GetReport.Value))/(3*stdev(GetReport.Value))) as CPK1, 
    --((avg(GetReport.Value)- GetReport.MinLimit)/(3*stdev(GetReport.Value))) as CPK2,
    Group_concat	(GetReport.Value, ', '),
	GetReport.SignalName,
	GetReport.CheckLimit ,
	Group_concat(GetReport.TimeStamp, ', '),
	Group_concat(GetReport.runid, ', ')
FROM
	(((planrun
	INNER JOIN testrun ON
		planrun.runid = testrun.planrunid)
	Inner JOIN GetScriptName ON
		testrun.runid = GetScriptName.runid)
	LEFT JOIN GetReport ON
		GetReport.runid = testrun.runid)
Group BY
	GetReport.ParentName , GetReport.StepName , GetReport.ValueName, GetReport.Verdict
--ORDER BY
--	 GetReport.ValueName
";
        private readonly string _statsBySerialSql = @"
-- Get Stats source
With GetReport AS 
	(
	WITH RECURSIVE GetStepName as (
		SELECT
			testrun.RunID , 
			TestRun.ParentRunID, 
			testRun.PlanRunID ,
			coalesce(params.Value, ' ') As ParentName,
			'' as StepName,
			0 as level
		FROM
			(testrun
		INNER JOIN testrun2params ON
			((testrun.RunID = testrun2params.Scope)
				AND (testrun.runid = testrun2params.runid))
		INNER JOIN params ON
				testrun2params.paramid = params.paramid)
		WHERE
			((((params.Name)= 'Name')
				AND ((params.GroupName)= '')))
			AND testrun.ParentRunID is NULL
		UNION
		SELECT
			e.RunID ,
			e.ParentRunID,
			e.PlanRunID,
			eh.ParentName || '\' || eh.StepName ,
			params.Value ,
			eh.level + 1 as level
		FROM
			testrun e
		join GetStepName eh ON
			e.ParentRunID = eh.RunID
		INNER JOIN testrun2params ON
			( (e.runid = testrun2params.runid )
				AND (e.runid = testrun2params.Scope))
		INNER JOIN params ON
				testrun2params.paramid = params.paramid
		WHERE
			((((params.Name)= 'Name')
				AND ((params.GroupName)= '')))
		), 
	  GetPlanRuns AS 
	  (
		  -- GetPLanRunsbySerialNumber
			-- returns GetPlanRuns.PlanRunNumber
	  WITH	 GetStepName AS (
		SELECT
			testrun2params.scope,
			testrun2params.runid,
			params.name,
			params.value AS StepName,
			params.groupname
		FROM
			(testrun
		INNER JOIN testrun2params ON
			testrun.runid = testrun2params.runid)
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((testrun2params.scope)= testrun.runid)
				And ((testrun2params.runid)= testrun.runid)
					And ((params.name)= 'Step Name')))
		,
		GetVerdict AS
			 (
			SELECT
				planrun.planrunnumber,
				params.value as Verdict
			FROM
				planrun
			INNER JOIN ((testrun
			INNER JOIN testrun2params ON
				testrun.runid = testrun2params.runid)
			INNER JOIN params ON
				testrun2params.paramid = params.paramid) ON
				planrun.runid = testrun.runid
			WHERE
				(((testrun2params.scope)= testrun.runid)
					And ((testrun2params.runid)= testrun.runid)
						And ((params.name)= 'Verdict'))
			),
			 GetSerialNumber AS
			 (
				SELECT
					testrun.runid,
					testrun.parentrunid,
					testrun2params.scope,
					params.Value AS SerialNumber
				FROM
					(testrun
				INNER JOIN testrun2params ON
					(testrun.ParentRunID = testrun2params.Scope)
						AND (testrun.runid = testrun2params.runid))
				INNER JOIN params ON
					testrun2params.paramid = params.paramid
				WHERE
					(((params.Name)= 'Name')
						AND ((params.GroupName)= ''))
			)
			SELECT
				planrun.planrunnumber,
				result.dim0 as Signal,
				result.dim1 as SerialNumber
			FROM
				(((((planrun
			INNER JOIN testrun ON
				planrun.runid = testrun.planrunid)
			LEFT JOIN resultseries ON
				testrun.runid = resultseries.runid)
			LEFT JOIN result ON
				resultseries.resultseriesid = result.resultseriesid)
			LEFT JOIN resulttype ON
				resultseries.resulttypeid = resulttype.resulttypeid)
			LEFT JOIN GetStepName ON
				testrun.runid = GetStepName.runid)
			WHERE
				( (GetStepName.StepName = 'User Input Test Info Step')
					and (Signal = 'Serial Number')
						and (SerialNumber =@serNum))
			ORDER BY
				planrun.PlanRunNumber
	),
	--Got plan runs
	 GetStepMinLimit AS
	(
	SELECT
		testrun2params.scope,
		testrun2params.runid,
		params.groupname,
		params.name,
		params.value AS MinLimit
	FROM
		(testrun
	INNER JOIN testrun2params ON
		testrun.runid = testrun2params.runid)
	INNER JOIN params ON
		testrun2params.paramid = params.paramid
	WHERE
		(((testrun2params.scope)= testrun.runid)
			And ((testrun2params.runid)= testrun.runid)
				And ((params.groupname)= 'Limits')
					And ((params.name)= 'Minimum Value')))
	,
	 GetStepMaxLimit AS 
	(
	SELECT
		testrun2params.scope,
		testrun2params.runid,
		params.name,
		params.value AS MaxLimit,
        params.groupname
	FROM
		(testrun
	INNER JOIN testrun2params ON
		testrun.runid = testrun2params.runid)
	INNER JOIN params ON
		testrun2params.paramid = params.paramid
	WHERE
		(((testrun2params.scope)= testrun.runid)
			And ((testrun2params.runid)= testrun.runid)
				And ((params.name)= 'Maximum Value')
					And ((params.groupname)= 'Limits')))
	,
	GetStepVerdict AS (
	SELECT
		testrun2params.scope,
		testrun2params.runid,
		params.name,
		params.value AS Verdict
	FROM
		(testrun
	INNER JOIN testrun2params ON
		testrun.runid = testrun2params.runid)
	INNER JOIN params ON
		testrun2params.paramid = params.paramid
	WHERE
		(((testrun2params.scope)= testrun.runid)
			And ((testrun2params.runid)= testrun.runid)
				And ((params.name)= 'Verdict')
					And ((params.value)<> 'NotSet')))
	,
	GetStepTimestamp AS (
	SELECT
		testrun2params.scope,
		testrun2params.runid,
		params.name,
		params.value AS [TimeStamp],
		params.groupname
	FROM
		(testrun
	INNER JOIN testrun2params ON
		testrun.runid = testrun2params.runid)
	INNER JOIN params ON
		testrun2params.paramid = params.paramid
	WHERE
		(((testrun2params.scope)= testrun.runid)
			And ((testrun2params.runid)= testrun.runid)
				And ((params.name)= 'StartTime')))
	,
	 GetStepCheckLimit AS (
	SELECT
		testrun2params.scope,
		testrun2params.runid,
		params.name,
		params.value AS CheckLimit,
		params.groupname
	FROM
		(testrun
	INNER JOIN testrun2params ON
		testrun.runid = testrun2params.runid)
	INNER JOIN params ON
		testrun2params.paramid = params.paramid
	WHERE
		(((testrun2params.scope)= testrun.runid)
			And ((testrun2params.runid)= testrun.runid)
				And ((params.name)= 'Check Limits')
					And ((params.groupname)= 'Limits')))
	SELECT
		GetPlanRuns.SerialNumber,
		GetStepName.ParentName,
		GetStepName.StepName,
		GetStepTimestamp.TimeStamp,
		resulttype.Dim0 as SignalName,
		Case
			WHEN 
				((resulttype.Dim0 = 'Measurement')
					OR (resulttype.Dim0 = 'Calibration')
					Or (resulttype.Dim0 = 'Source')
					OR (resulttype.Dim0 = 'Signal')
					OR (resulttype.Dim0 = 'Value')
					)
				THEN
					COALESCE(result.Dim0, 'Result.dim0' )----cast (result.Dim0 as TEXT)
			WHEN
			   (resulttype.Dim0 = 'IPAddress')
				THEN 
				    'IPADR' ----cast ('IPADDRESS' as TEXT)
            ELSE
				COALESCE(resulttype.Dim0, 'NULL')
		END ValueName,
		result.dim0 as Signal,
		--resulttype.Dim1 as ValueName,
		--result.dim1 as Value,
		Case
			WHEN 
				((resulttype.Dim0 = 'Measurement')
					OR (resulttype.Dim0 = 'Calibration')
					Or (resulttype.Dim0 = 'Source')
					OR (resulttype.Dim0 = 'Signal')
					OR (resulttype.Dim0 = 'Value')
					)
				THEN
					cast (result.dim1  as REAL)
			WHEN
			   (resulttype.Dim0 = 'IPAddress')
				THEN 
				    cast (0.0   as REAL)
            WHEN ((resulttype.DIM0 = NULL)
                 OR (result.Dim0 = NULL )) 
                THEN 
                  cast (0.0   as REAL)
			ELSE
				cast (COALESCE(result.Dim0, 0.0) as REAL) --cast (result.dim0    as REAL)
		END Value,
		COALESCE(GetStepMinLimit.MinLimit, 0.0)  as MinLimit,
		COALESCE(GetStepMaxLimit.MaxLimit, 0.0)  as MaxLimit,
		GetStepVerdict.Verdict,
		GetStepCheckLimit.CheckLimit,
		resulttype.Dim2 as ValueName2,
		result.Dim2 as Value2,
		resulttype.Dim3 as ValueName3,
		result.Dim3 as Value3,
		planrun.planrunnumber,
		testrun.runid
	FROM
		(((((((((((
		GetPlanRuns
		LEFT Join planrun ON
			planrun.PlanRunNumber = GetPlanRuns.PlanRunNumber)
		LEFT JOIN testrun ON
			planrun.runid = testrun.planrunid)
		LEFT JOIN resultseries ON
			testrun.runid = resultseries.runid)
		LEFT JOIN result ON
			resultseries.resultseriesid = result.resultseriesid)
		LEFT JOIN resulttype ON
			resultseries.resulttypeid = resulttype.resulttypeid)
		LEFT JOIN GetStepName ON
			testrun.runid = GetStepName.runid)
		LEFT JOIN GetStepMinLimit ON
			testrun.runid = GetStepMinLimit.runid)
		LEFT JOIN GetStepMaxLimit ON
			testrun.runid = GetStepMaxLimit.runid)
		LEFT JOIN GetStepVerdict ON
			testrun.runid = GetStepVerdict.runid)
		LEFT JOIN GetStepTimestamp ON
			testrun.runid = GetStepTimestamp.runid)
		LEFT JOIN GetStepCheckLimit ON
			testrun.RunID = GetStepCheckLimit.runid)
	WHERE
		((SignalName = 'Measurement')
			OR (SignalName = 'Calibration')
			Or (SignalName = 'Source')
			--AND (CheckLimit = TRUE)
			)
	ORDER BY
		StepName
),

	 GetScriptName AS
		 (
		SELECT
			testrun.planrunid,
			testrun.runid,
			testrun.parentrunid,
			testrun2params.scope,
			params.Value AS ScriptName
		FROM
			(testrun
		INNER JOIN testrun2params ON
			(testrun.PlanRunID = testrun2params.Scope)
				AND (testrun.runid = testrun2params.runid))
		INNER JOIN params ON
			testrun2params.paramid = params.paramid
		WHERE
			(((params.Name)= 'Name')
				AND ((params.GroupName)= ''))
		)
Select
	COALESCE(GetScriptName.ScriptName, ' ') as Scriptname, 
	COALESCE(GetReport.SerialNumber, ' ') as SerialNumber,
	COALESCE(GetReport.ParentName, ' ') as ParentName, 
	COALESCE(GetReport.StepName, ' ') as StepName,
	COALESCE(GetReport.ValueName, ' ') as ValueName,
	COALESCE(GetReport.Signal, ' ') as Signal,
	Verdict,
	COALESCE(avg(GetReport.Value),0.0) as Average,
	cast (COALESCE(Min(GetReport.Value),0.0) as real) as Minimum,
	cast (COALESCE(Max(GetReport.Value),0.0) as real) as Maximum,
	COALESCE(variance(GetReport.Value), 0.0) as Variance,
	COALESCE(stdev(GetReport.Value),0.0) as StandardDeviation,
	Count(GetReport.Value) as Count,
	COALESCE(GetReport.MinLimit, 0.0)  as MinLimit,
	COALESCE(GetReport.MaxLimit, 0.0)  as MaxLimit,
	---GetReport.MinLimit,
	---GetReport.MaxLimit,
	--min((GetReport.MaxLimit - avg(GetReport.Value))/(3*stdev(GetReport.Value)), (avg(GetReport.Value)- GetReport.MinLimit)/(3*stdev(GetReport.Value))) as CPK,
	--((GetReport.MaxLimit - avg(GetReport.Value))/(3*stdev(GetReport.Value))) as CPK1, 
    --((avg(GetReport.Value)- GetReport.MinLimit)/(3*stdev(GetReport.Value))) as CPK2,
    ---Group_concat	(GetReport.Value, ', '),
	GetReport.SignalName,
	GetReport.CheckLimit 
	--Group_concat(GetReport.TimeStamp, ', '),
	--Group_concat(GetReport.runid, ', ')
FROM
	(((planrun
	INNER JOIN testrun ON
		planrun.runid = testrun.planrunid)
	Inner JOIN GetScriptName ON
		testrun.runid = GetScriptName.runid)
	LEFT JOIN GetReport ON
		GetReport.runid = testrun.runid)
Group BY
	GetReport.ParentName , GetReport.StepName , GetReport.ValueName, GetReport.Verdict
--ORDER BY
--	 GetReport.ValueName
";
		
	public MainForm()
        {
            InitializeComponent(); dgvOverview.AutoGenerateColumns = true;
            dgvDetails.AutoGenerateColumns = true; 
            dgvStatsAll.AutoGenerateColumns = true;
            dgvStatsSerial.AutoGenerateColumns = true;
            cboVerdict.Items.AddRange(new object[] { "(Any)", "Pass", "Fail", "NotSet" });
            cboVerdict.SelectedIndex = 0;
            cboStatus.Items.Add("(Any)");
            cboStatus.SelectedIndex = 0; 
            dtFrom.ShowCheckBox = true;
            dtFrom.Value = new DateTime(2015,1,1);
            dtTo.ShowCheckBox = true; LoadSettings();
            UpdateUiState();
        }
        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e); if (File.Exists(_dbPath)) { await LoadOverviewAsync(); }
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSettings(); base.OnFormClosing(e);
        }
        private struct VarianceAccumulator
        {
            public long Count;
            public double Mean;
            public double M2; // Sum of squares of differences
        }
        private static void RegisterVariance(SqliteConnection conn)
        {
            conn.CreateAggregate<double?, VarianceAccumulator, double?>(
                name: "variance",
                seed: new VarianceAccumulator
                {
                    Count = 0,
                    Mean = 0.0,
                    M2 = 0.0
                },
                func: (acc, value) =>
                {
                    if (value.HasValue)
                    {
                        acc.Count++;
                        var delta = value.Value - acc.Mean;
                        acc.Mean += delta / acc.Count;
                        acc.M2 += delta * (value.Value - acc.Mean);
                    }
                    return acc;
                },
                resultSelector: acc =>
                    acc.Count > 1
                        ? acc.M2 / (acc.Count - 1)   // ✅ sample variance
                        : (double?) 0,
                isDeterministic: true
            );
		
		
        }

        private static void RegisterStdDev(SqliteConnection conn)
        {
            conn.CreateAggregate<double?, VarianceAccumulator, double?>(
                name: "stdev",
                seed: new VarianceAccumulator
                {
                    Count = 0,
                    Mean = 0.0,
                    M2 = 0.0
                },
                func: (acc, value) =>
                {
                    if (value.HasValue)
                    {
                        acc.Count++;
                        var delta = value.Value - acc.Mean;
                        acc.Mean += delta / acc.Count;
                        acc.M2 += delta * (value.Value - acc.Mean);
                    }
                    return acc;
                },
                resultSelector: acc =>
                    acc.Count > 1
                        ? (double?)Math.Sqrt(acc.M2 / (acc.Count - 1))    // ✅ sample stdev
                        : (double?)0.0,
                isDeterministic: true
            );
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("lastDbPath", out var el))
                    {
                        var p = el.GetString();
                        if (!string.IsNullOrWhiteSpace(p) && File.Exists(p))
                        {
                            _dbPath = p!; txtDbPath.Text = _dbPath;
                        }
                    }
                }
            }
            catch { }
        }
        private void SaveSettings() { try { Directory.CreateDirectory(SettingsFolder); var json = JsonSerializer.Serialize(new { lastDbPath = _dbPath }, new JsonSerializerOptions { WriteIndented = true }); File.WriteAllText(SettingsFile, json); } catch { } }
        private void UpdateUiState() { bool hasDb = File.Exists(_dbPath); btnLoad.Enabled = hasDb; btnLoadStatsAll.Enabled = hasDb; btnLoadStatsSerial.Enabled = hasDb && !string.IsNullOrWhiteSpace(txtSerial.Text); btnExportOverview.Enabled = dgvOverview.DataSource is DataTable dtO && dtO.Rows.Count > 0; btnExportDetails.Enabled = dgvDetails.DataSource is DataTable dtD && dtD.Rows.Count > 0; btnExportStatsAll.Enabled = dgvStatsAll.DataSource is DataTable dtA && dtA.Rows.Count > 0; btnExportStatsSerial.Enabled = dgvStatsSerial.DataSource is DataTable dtS && dtS.Rows.Count > 0; }
        private struct StatAcc { public long Count; public double Mean; public double M2; }
        private static bool _udfRegistered = false;
        private static void RegisterCustomFunctions(SqliteConnection conn) 
        {
            RegisterVariance(conn);
            RegisterStdDev(conn);
            if (_udfRegistered) return;

            //            conn.CreateAggregate<double?, StatAcc, double?>("variance", new StatAcc { Count = 0, Mean = 0.0, M2 = 0.0 }, (acc, x) => { if (x.HasValue) { acc.Count++; var delta = x.Value - acc.Mean; acc.Mean += delta / acc.Count; var delta2 = x.Value - acc.Mean; acc.M2 += delta * delta2; } return acc; }, acc => acc.Count > 1 ? (double?)(acc.M2 / (acc.Count - 1)) : null, true); conn.CreateAggregate<double?, StatAcc, double?>("stdev", new StatAcc { Count = 0, Mean = 0.0, M2 = 0.0 }, (acc, x) => { if (x.HasValue) { acc.Count++; var delta = x.Value - acc.Mean; acc.Mean += delta / acc.Count; var delta2 = x.Value - acc.Mean; acc.M2 += delta * delta2; } return acc; }, acc => acc.Count > 1 ? (double?)Math.Sqrt(acc.M2 / (acc.Count - 1)) : null, true); 
            _udfRegistered = true; 
        }
        private async Task<DataTable> ExecuteQueryAsync(string sql, Action<SqliteCommand> configure) 
        { 
            DataTable table = new DataTable(); 
            using var conn = new SqliteConnection(ConnectionString);

         //   SQLiteFunction.RegisterFunction(
         //       typeof(TestRunViewerSqlite_Pro.Data.SqliteFunctions.VarianceAggregate));

            await conn.OpenAsync(); RegisterCustomFunctions(conn); 
            using var cmd = new SqliteCommand(sql, conn); 
            configure?.Invoke(cmd); 
            //cmd.
            using var reader = await cmd.ExecuteReaderAsync();

            table.BeginLoadData();
            table.Constraints.Clear();
            table.PrimaryKey = null;    

            table.Load(reader); 
            table.EndLoadData();

            return table;
        }
        private async Task<DataTable> ExecuteDetailQueryAsync(string sql, Action<SqliteCommand> configure)
        {
            DataTable table = new DataTable();
            using var conn = new SqliteConnection(ConnectionString);

            //   SQLiteFunction.RegisterFunction(
            //       typeof(TestRunViewerSqlite_Pro.Data.SqliteFunctions.VarianceAggregate));

            await conn.OpenAsync(); RegisterCustomFunctions(conn);
            using var cmd = new SqliteCommand(sql, conn);
            configure?.Invoke(cmd);
            //cmd.
            using var reader = await cmd.ExecuteReaderAsync();

            //            table.PrimaryKey = null;
            //table.PrimaryKey = [1, 2, 3, 4, 5];
            /*
                        // 1️⃣ Define schema FIRST
                        table.Columns.Add("RunID", typeof(long));
                        table.Columns.Add("TimeStamp", typeof(string));
                        table.Columns.Add("Signal", typeof(string));
                        table.Columns.Add("Value", typeof(double));

                        // 2️⃣ Set composite PrimaryKey BEFORE loading
                        table.PrimaryKey = new[]
                        {
                            table.Columns["RunID"],
                            table.Columns["TimeStamp"],
                            table.Columns["Signal"]
                        };
            */
            table.BeginLoadData();
            table.Constraints.Clear();
            table.Columns.Add("PrimaryKey", typeof(string));
            table.PrimaryKey = new[]
            {
                            table.Columns["PrimaryKey"]
            };
            //table.PrimaryKey = null;

            table.Load(reader);
            table.EndLoadData();

            return table;
        }
        private string BuildFilteredOverviewWrapperSql()
            {
                return "SELECT * FROM (" + _overviewSql + ") ov WHERE 1=1\n{FILTERS}\nORDER BY ov.PlanRunNumber, ov.TimeStamp";
            }
       
        private string BuildFilteredOverviewWrapperSqlSN() { return "SELECT * FROM (" + _overviewSqlSN + ") ov WHERE 1=1\n{FILTERS}\nORDER BY ov.PlanRunNumber, ov.TimeStamp"; }

        private async Task PopulateStatusFilterAsync()
        {
            try
            {
                const string statusSql = @"
SELECT DISTINCT Status
FROM ({0})
WHERE Status IS NOT NULL
ORDER BY Status";

                var sql = string.Format(statusSql, _overviewSql);
                var table = await ExecuteQueryAsync(sql, null);

                cboStatus.Items.Clear();
                cboStatus.Items.Add("(Any)");

                foreach (DataRow row in table.Rows)
                {
                    cboStatus.Items.Add(Convert.ToString(row[0]) ?? string.Empty);
                }

                cboStatus.SelectedIndex = 0;
                lblStatusFilter.Visible = cboStatus.Visible = true;
            }
            catch
            {
                lblStatusFilter.Visible = cboStatus.Visible = false;
            }
        }
        private void AddOverviewFilterParameters(SqliteCommand cmd) 
        { 
            cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", " {FILTERS} "); 
            if (!string.IsNullOrWhiteSpace(txtPlanRunName.Text)) 
            { 
                cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.PlanRunName LIKE @name\n{FILTERS}"); 
                cmd.Parameters.Add(new SqliteParameter("@name", "%" + txtPlanRunName.Text.Trim() + "%")); 
            } 
            if (cboVerdict.SelectedIndex > 0) 
            { 
                cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.Verdict = @verdict\n{FILTERS}"); 
                cmd.Parameters.Add(new SqliteParameter("@verdict", Convert.ToString(cboVerdict.SelectedItem))); 
            } 
            if (cboStatus.Visible && cboStatus.SelectedIndex > 0) 
            { 
                cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.Status = @status\n{FILTERS}"); 
                cmd.Parameters.Add(new SqliteParameter("@status", Convert.ToString(cboStatus.SelectedItem))); 
            } 
            if (dtFrom.Checked) 
            { 
                cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.TimeStamp >= @from\n{FILTERS}"); 
                cmd.Parameters.Add(new SqliteParameter("@from", dtFrom.Value.ToString("yyyy-MM-dd HH:mm:ss"))); 
            } 
            if (dtTo.Checked) 
            { 
                cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", "AND ov.TimeStamp <= @to\n{FILTERS}"); 
                cmd.Parameters.Add(new SqliteParameter("@to", dtTo.Value.ToString("yyyy-MM-dd HH:mm:ss"))); 
            } 
            cmd.CommandText = cmd.CommandText.Replace("{FILTERS}", string.Empty); }
/*
        private void AddOverviewFilterParameters(SqliteCommand cmd)
        {
            var sql = new StringBuilder(
                cmd.CommandText.Replace("{FILTERS}", string.Empty)
            );

            if (!string.IsNullOrWhiteSpace(txtPlanRunName.Text))
            {
                sql.AppendLine("AND ov.PlanRunName LIKE @name");
                cmd.Parameters.AddWithValue(
                    "@name",
                    $"%{txtPlanRunName.Text.Trim()}%");
            }

            if (cboVerdict.SelectedIndex > 0)
            {
                sql.AppendLine("AND ov.Verdict = @verdict");
                cmd.Parameters.AddWithValue(
                    "@verdict",
                    cboVerdict.SelectedItem?.ToString());
            }

            if (cboStatus.Visible && cboStatus.SelectedIndex > 0)
            {
                sql.AppendLine("AND ov.Status = @status");
                cmd.Parameters.AddWithValue(
                    "@status",
                    cboStatus.SelectedItem?.ToString());
            }

            if (dtFrom.Checked)
            {
                sql.AppendLine("AND ov.TimeStamp >= @from");
                cmd.Parameters.AddWithValue(
                    "@from",
                    dtFrom.Value.ToString("yyyy-MM-dd HH:mm:ss"));
            }

            if (dtTo.Checked)
            {
                sql.AppendLine("AND ov.TimeStamp <= @to");
                cmd.Parameters.AddWithValue(
                    "@to",
                    dtTo.Value.ToString("yyyy-MM-dd HH:mm:ss"));
            }

            cmd.CommandText = sql.ToString();
        }
*/
        private async Task LoadOverviewAsync()
        {
            if (!File.Exists(_dbPath))
            {
                MessageBox.Show(
                    this,
                    "Please open a SQLite database first.",
                    "No Database",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            ToggleUi(false);
            lblStatus.Text = "Loading overview...";

            try
            {
                await PopulateStatusFilterAsync();
				var sql =  (chkWithSerialNumber.Checked) ?
	                BuildFilteredOverviewWrapperSqlSN() : BuildFilteredOverviewWrapperSql();
				var table = await ExecuteQueryAsync(sql, AddOverviewFilterParameters);

                dgvOverview.DataSource = table;
                lblStatus.Text = $"Loaded {table.Rows.Count} rows.";

                PopulatePlanRunNumberMapping(table);

                if (table.Rows.Count > 0)
                {
                    dgvOverview.ClearSelection();
                    dgvOverview.Rows[0].Selected = true;
                    dgvDetails.DataSource = null;

                    await LoadSelectedSummaryAsync();
                    //await LoadSelectedDetailsAsync();
                }
                else
                {
                    dgvDetails.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Error Loading Overview",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblStatus.Text = "Failed to load overview.";
            }
            finally
            {
                ToggleUi(true);
                UpdateUiState();
            }
        }

        private void PopulatePlanRunNumberMapping(DataTable table)
        {
            cmbMapPlanRunNumber.Items.Clear();

            foreach (DataColumn column in table.Columns)
            {
                cmbMapPlanRunNumber.Items.Add(column.ColumnName);
            }

            var defaultColumn =
                table.Columns
                     .Cast<DataColumn>()
                     .Select(c => c.ColumnName)
                     .FirstOrDefault(
                         n => string.Equals(
                             n,
                             "PlanRunNumber",
                             StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(defaultColumn))
            {
                cmbMapPlanRunNumber.SelectedItem = defaultColumn;
            }
            else if (cmbMapPlanRunNumber.Items.Count > 0)
            {
                cmbMapPlanRunNumber.SelectedIndex = 0;
            }
        }
        private object? GetSelectedOverviewCell(string columnName)
        {
            if (dgvOverview.CurrentRow?.DataBoundItem is not DataRowView view)
                return null;

            var row = view.Row;

            return row.Table.Columns.Contains(columnName)
                ? row[columnName]
                : null;
        }

        private async Task LoadSelectedSummaryAsync()
        {
            if (dgvOverview.CurrentRow == null)
            {
                dgvDetails.DataSource = null;
                return;
            }

            if (cmbMapPlanRunNumber.SelectedItem is null)
            {
                lblStatus.Text = "Select the PlanRunNumber column mapping first.";
                return;
            }

            var columnName = cmbMapPlanRunNumber.SelectedItem.ToString();
            var rawValue = GetSelectedOverviewCell(columnName!);
			var sercolname = "SerialNumber";
			var SerialNmbr = GetSelectedOverviewCell(sercolname!);
			txtSerial.Text = $"{SerialNmbr}";
            if (rawValue is null || rawValue == DBNull.Value)
            {
                lblStatus.Text = $"No value in column '{columnName}'.";
                dgvDetails.DataSource = null;
                return;
            }

            if (!long.TryParse(Convert.ToString(rawValue), out var planRunNumber))
            {
                lblStatus.Text =
                    $"Selected PlanRunNumber ('{columnName}') is not numeric.";
                dgvDetails.DataSource = null;
                return;
            }

            ToggleUi(false);
            lblStatus.Text =
                $"Loading Summary for PlanRunNumber {planRunNumber}...";

            try
            {


                var qstr = string.Format($"{_summarySql} {planRunNumber});");
                var table = await ExecuteQueryAsync(
                    //qstr, //null); //, //
                    _summarySql,
                    cmd => cmd.Parameters.AddWithValue(
                        "@PlanRunNumber",
                        planRunNumber));

                dgvSummary.DataSource = table;
                lblStatus.Text =
                    $"Loaded {table.Rows.Count} summary rows for PlanRunNumber {planRunNumber}.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Error Loading Summary",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblStatus.Text = "Failed to load Summary.";
            }
            finally
            {
                ToggleUi(true);
                UpdateUiState();
            }
        }

        private async Task LoadSelectedDetailsAsync()
        {
            if (dgvOverview.CurrentRow == null)
            {
                dgvDetails.DataSource = null;
                return;
            }

            if (cmbMapPlanRunNumber.SelectedItem is null)
            {
                lblStatus.Text = "Select the PlanRunNumber column mapping first.";
                return;
            }

            var columnName = cmbMapPlanRunNumber.SelectedItem.ToString();
            var rawValue = GetSelectedOverviewCell(columnName!);

            if (rawValue is null || rawValue == DBNull.Value)
            {
                lblStatus.Text = $"No value in column '{columnName}'.";
                dgvDetails.DataSource = null;
                return;
            }

            if (!long.TryParse(Convert.ToString(rawValue), out var planRunNumber))
            {
                lblStatus.Text =
                    $"Selected PlanRunNumber ('{columnName}') is not numeric.";
                dgvDetails.DataSource = null;
                return;
            }

            ToggleUi(false);
            lblStatus.Text =
                $"Loading details for PlanRunNumber {planRunNumber}...";

            try
            {
                

                var qstr = string.Format($"{_detailsSql} {planRunNumber});");
                var table = await ExecuteDetailQueryAsync(
                    //qstr, //null); //, //
                    _detailsSql, 
                    cmd => cmd.Parameters.AddWithValue(
                        "@PlanRunNumber",
                        planRunNumber));

                dgvDetails.DataSource = table;
                lblStatus.Text =
                    $"Loaded {table.Rows.Count} detail rows for PlanRunNumber {planRunNumber}.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Error Loading Details",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblStatus.Text = "Failed to load details.";
            }
            finally
            {
                ToggleUi(true);
                UpdateUiState();
            }
        }

        private async Task LoadStatsAllAsync()
        {
            if (!File.Exists(_dbPath))
                return;

            lblStatus.Text = "Loading stats (all runs)...";
            ToggleUi(true);
            UpdateUiState();
            ToggleUi(false);
            
            try
            {
                var table = await ExecuteQueryAsync(_statsAllSql, null);
                dgvStatsAll.DataSource = table;
                lblStatus.Text = $"Loaded {table.Rows.Count} rows.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Error Loading Stats (All)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblStatus.Text = "Failed to load stats (all).";
            }
            finally
            {
                ToggleUi(true);
                UpdateUiState();
            }
        }

        private async Task LoadStatsBySerialAsync(string serial)
        {
            if (!File.Exists(_dbPath))
                return;

            ToggleUi(false);
            lblStatus.Text = $"Loading stats (serial {serial})...";

            try
            {
                var table = await ExecuteQueryAsync(
                    _statsBySerialSql,
                    cmd => cmd.Parameters.AddWithValue("@serNum", serial));
				if (table.Rows.Count > 1)
				{
					dgvStatsSerial.DataSource = table;
					lblStatus.Text = $"Loaded {table.Rows.Count} rows.";
				}
				else lblStatus.Text = "Loaded 0 rows. More than one run needed for statistical analysis! ";
            }

            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Error Loading Stats (By Serial)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblStatus.Text = "Failed to load stats (serial).";
            }
            finally
            {
                ToggleUi(true);
                UpdateUiState();
            }
        }

        private void ExportDataTableToCsv(DataTable table, string path)
        {
            using var writer = new StreamWriter(path, false, Encoding.UTF8);

            // Header
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (i > 0)
                    writer.Write(',');

                writer.Write(
                    $"\"{table.Columns[i].ColumnName.Replace("\"", "\"\"")}\"");
            }

            writer.WriteLine();

            // Rows
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (i > 0)
                        writer.Write(',');

                    var value =
                        row[i]?.ToString()?.Replace("\"", "\"\"") ?? string.Empty;

                    writer.Write($"\"{value}\"");
                }

                writer.WriteLine();
            }
        }
        private void ExportDataTableToExcel(DataTable table, string path, string sheetName) 
        { 
            using var wb = new XLWorkbook();
			var cellstring = "";
            var ws = wb.Worksheets.Add(string.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName); 
            for (int c = 0; c < table.Columns.Count; c++) 
                ws.Cell(1, c + 1).Value = table.Columns[c].ColumnName;
			for (int r = 0; r < table.Rows.Count; r++)
				for (int c = 0; c < table.Columns.Count; c++)
				{//var value = table.Rows[r][c];
				 //ws.Cell(r + 2, c + 1).SetValue(value?.ToString() ?? string.Empty);
					cellstring = table.Rows[r][c]?.ToString() ?? string.Empty;
					//ws.Cell(r + 2, c + 1).SetValue(table.Rows[r][c]?.ToString() ?? string.Empty); 
					ws.Cell(r + 2, c + 1).SetValue(cellstring.Length >32767 ? cellstring.Substring(0,32767) : cellstring);
					//ws.Cell(r + 2, c + 1).Value = table.Rows[r][c]; 
				}
            ws.Columns().AdjustToContents(); wb.SaveAs(path); }
        private void Export(DataTable dt, string defaultFile)
        {
            if (dt == null || dt.Rows.Count == 0)
                return;

            using var saveDialog = new SaveFileDialog
            {
                Title = "Export",
                Filter =
                    "Excel Workbook (*.xlsx)|*.xlsx|" +
                    "CSV Files (*.csv)|*.csv|" +
                    "All files (*.*)|*.*",
                FileName = defaultFile
            };

            if (saveDialog.ShowDialog(this) != DialogResult.OK)
                return;

            try
            {
                var extension = Path
                    .GetExtension(saveDialog.FileName)
                    .ToLowerInvariant();

                if (extension == ".xlsx")
                {
                    ExportDataTableToExcel(
                        dt,
                        saveDialog.FileName,
                        Path.GetFileNameWithoutExtension(defaultFile));
                }
                else
                {
                    ExportDataTableToCsv(dt, saveDialog.FileName);
                }

                MessageBox.Show(
                    this,
                    "Exported.",
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void ToggleUi(bool enabled)
        {
            btnOpenDb.Enabled = enabled;
            btnLoad.Enabled = enabled && File.Exists(_dbPath);

            dgvOverview.Enabled = enabled;
            dgvDetails.Enabled = enabled;
            dgvStatsAll.Enabled = enabled;
            dgvStatsSerial.Enabled = enabled;

            grpFilters.Enabled = enabled;
            grpMapping.Enabled = enabled;
            grpStats.Enabled = enabled;

            btnExportOverview.Enabled =
                enabled &&
                dgvOverview.DataSource is DataTable dtOverview &&
                dtOverview.Rows.Count > 0;

            btnExportDetails.Enabled =
                enabled &&
                dgvDetails.DataSource is DataTable dtDetails &&
                dtDetails.Rows.Count > 0;

            btnExportStatsAll.Enabled =
                enabled &&
                dgvStatsAll.DataSource is DataTable dtStatsAll &&
                dtStatsAll.Rows.Count > 0;

            btnExportStatsSerial.Enabled =
                enabled &&
                dgvStatsSerial.DataSource is DataTable dtStatsSerial &&
                dtStatsSerial.Rows.Count > 0;
        }

        private async void btnLoad_Click(object sender, EventArgs e)
        {
            await LoadOverviewAsync();
        }

        private async void dgvOverview_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvOverview.Focused || dgvOverview.IsHandleCreated)
            {
                //await LoadSelectedDetailsAsync();
                dgvDetails.DataSource = null;
                await LoadSelectedSummaryAsync();
            }
        }
        private async void btnGetDetails_Click(object sender, EventArgs e)
        {
            if (dgvOverview.Focused || dgvOverview.IsHandleCreated)
            {
                await LoadSelectedDetailsAsync();
                await LoadSelectedSummaryAsync();
            }
        }
        private async void btnApplyFilters_Click(object sender, EventArgs e)
        {
            await LoadOverviewAsync();
        }

        private async void btnLoadStatsAll_Click(object sender, EventArgs e)
        {
            await LoadStatsAllAsync();
        }

        private async void btnLoadStatsSerial_Click(object sender, EventArgs e)
        {
            var serial = txtSerial.Text?.Trim();

            if (!string.IsNullOrEmpty(serial))
            {
                await LoadStatsBySerialAsync(serial);
            }
            else
            {
                MessageBox.Show(
                    this,
                    "Enter a serial number.",
                    "Missing Serial",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }
        private void btnOpenDb_Click(object sender, EventArgs e)
        {
            using var openDialog = new OpenFileDialog
            {
                Title = "Open Test Results DB",
                Filter =
                    "Test Results DB (*.db;*.TapResults)|*.db;*.TapResults|" +
                    "All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (openDialog.ShowDialog(this) == DialogResult.OK)
            {
                _dbPath = openDialog.FileName;
                txtDbPath.Text = _dbPath;

                SaveSettings();
                LoadOverviewAsync();
                UpdateUiState();
            }
        }

        private void btnExportOverview_Click(object sender, EventArgs e)
        {
            if (dgvOverview.DataSource is DataTable table)
            {
                Export(table, "overview.xlsx");
            }
        }
        private void btnExportDetails_Click(object sender, EventArgs e)
        {
            if (dgvDetails.DataSource is DataTable table)
            {
                Export(table, "details.xlsx");
            }
        }

        private void btnExportStatsAll_Click(object sender, EventArgs e)
        {
            if (dgvStatsAll.DataSource is DataTable table)
            {
                Export(table, "stats_all.xlsx");
            }
        }

        private void btnExportStatsSerial_Click(object sender, EventArgs e)
        {
            if (dgvStatsSerial.DataSource is DataTable table)
            {
                Export(table, "stats_serial.xlsx");
            }
        }
    }
}