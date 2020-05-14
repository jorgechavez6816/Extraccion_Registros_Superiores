Sub Main
	Call TopNExtraction()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Datos: Extracción de registros superiores
Function TopNExtraction
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.TopRecordsExtraction
	task.IncludeAllFields
	task.AddKey "COD_PROD", "A"
	task.AddKey "TOTAL", "D"
	dbName = "Extraccion_Reg_Sup_01.IMD"
	task.OutputFileName = dbName
	task.NumberOfRecordsToExtract = 5
	task.CreateVirtualDatabase = False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function