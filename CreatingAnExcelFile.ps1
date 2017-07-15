<#
.SYNOPSIS
    This script copies a .xls file written using a different version than the one currently on your computer and updates the copy.
 
.DESCRIPTION
    This script copies a .xls file written using a different version than the one currently on your computer and updates the copy.
    The purpose of this script is to show you how you can bring the "Microsoft Excel - Compatibility Checker" window that pops up, to the front automatically.
 
.INPUTS
    H:\My Documents\FileToCopyAndUpdate.xls
 
.OUTPUTS
    H:\My Documents\UpdatedCopyOfOriginalFile.xls
 
.NOTES
    Author: dklempfner@gmail.com
    Date: 14/07/2017
#>
 
function BringExcelCompatibilityCheckerToFront
{
    $TypeDef1 = @"
  using System;
  using System.Runtime.InteropServices;
  public class Tricks {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@
    $TypeDef2 = @"
 
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
 
namespace Api
{
 
public class WinStruct
{
   public string WinTitle {get; set; }
   public int MainWindowHandle { get; set; }
}
 
public class ApiDef
{
   private delegate bool CallBackPtr(int hwnd, int lParam);
   private static CallBackPtr callBackPtr = Callback;
   private static List<WinStruct> _WinStructList = new List<WinStruct>();
 
   [DllImport("User32.dll")]
   [return: MarshalAs(UnmanagedType.Bool)]
   private static extern bool EnumWindows(CallBackPtr lpEnumFunc, IntPtr lParam);
 
   [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
 
   private static bool Callback(int hWnd, int lparam)
   {
       StringBuilder sb = new StringBuilder(256);
       int res = GetWindowText((IntPtr)hWnd, sb, 256);
      _WinStructList.Add(new WinStruct { MainWindowHandle = hWnd, WinTitle = sb.ToString() });
       return true;
   }  
 
   public static List<WinStruct> GetWindows()
   {
      _WinStructList = new List<WinStruct>();
      EnumWindows(callBackPtr, IntPtr.Zero);
      return _WinStructList;
   }
 
}
}
"@
 
    Add-Type -TypeDefinition $TypeDef1 -Language CSharpVersion3
    Add-Type -TypeDefinition $TypeDef2 -Language CSharpVersion3
 
    $excelInstance = $null
 
    do
    {
        $excelInstance = [Api.Apidef]::GetWindows() | Where-Object { $_.WinTitle.ToUpper() -eq "Microsoft Excel - Compatibility Checker".ToUpper() }
    }
    while($null -eq $excelInstance)
   
    if($excelInstance.MainWindowHandle)
    {
        [void][Tricks]::SetForegroundWindow($excelInstance.MainWindowHandle)
    }
}
 
$excelApp = New-Object -ComObject Excel.Application
$excelApp.DisplayAlerts = $false
   
$xlsx = $excelApp.Workbooks.Open('H:\My Documents\FileToCopyAndUpdate.xls')
$worksheetObject = $Xlsx.Worksheets.Add()
$worksheet = $Xlsx.Worksheets.Item($worksheetObject.Index)   
$worksheet.Name =  'First worksheet'
$Row = 1
$column = 1
$worksheet.Cells.Item($Row, $column) = 'this is the value for cell A1'
 
#The SaveAs() function below opens up the "Microsoft Excel - Compatibility Checker" window.
#The main thread is blocked until you click "OK".
#The BringExcelCompatibilityCheckerToFront() function brings the window to the front.
#This must be done in a different thread since the main thead is blocked.
 
$job = Start-Job { BringExcelCompatibilityCheckerToFront }
$xlsx.SaveAs('H:\My Documents\UpdatedCopyOfOriginalFile.xls')
$excelApp.Quit()