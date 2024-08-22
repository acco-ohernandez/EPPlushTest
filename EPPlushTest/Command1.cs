using OfficeOpenXml;
namespace EPPlushTest
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            return Result.Succeeded;
        }

        public void CreateExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or LicenseContext.Commercial

            var file = new FileInfo(@"C:\path\to\your\file.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].Value = "Hello, EPPlus!";
                package.Save();
            }
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";
            string? methodBase = MethodBase.GetCurrentMethod().DeclaringType?.FullName;

            if (methodBase == null)
            {
                throw new InvalidOperationException("MethodBase.GetCurrentMethod().DeclaringType?.FullName is null");
            }
            else
            {
                Common.ButtonDataClass myButtonData1 = new Common.ButtonDataClass(
                    buttonInternalName,
                    buttonTitle,
                    methodBase,
                    Properties.Resources.Blue_32,
                    Properties.Resources.Blue_16,
                    "This is a tooltip for Button 1");

                return myButtonData1.Data;
            }
        }
    }

}
