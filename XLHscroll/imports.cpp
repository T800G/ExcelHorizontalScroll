//////uncomment to generate tlh files
//////#import "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL" named_guids raw_interfaces_only \
////	rename("RGB", "MSORGB")
////	rename("DocumentProperties", "MSODocumentProperties")
//////using namespace Office;
//////#import "C:\\Program Files\\Common Files\\microsoft shared\\VBA\\VBA6\\VBE6EXT.OLB" named_guids raw_interfaces_only
//////#import "C:\\Program Files\\Common Files\\DESIGNER\\MSADDNDR.DLL" named_guids raw_interfaces_only
////#import "C:\\Program Files\\Microsoft Office\\Office11\\EXCEL.EXE"  /*named_guids*/ raw_interfaces_only \
////    rename("DialogBox", "ExcelDialogBox") \
////    rename("RGB", "ExcelRGB") \
////    rename("CopyFile", "ExcelCopyFile") \
////    rename("ReplaceText", "ExcelReplaceText") \
////	//no_auto_exclude
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" \
    rename("RGB", "MSORGB") \
    rename("DocumentProperties", "MSODocumentProperties")
//using namespace Office;

#import "libid:0002E157-0000-0000-C000-000000000046"
//using namespace VBIDE;

#import "libid:00020813-0000-0000-C000-000000000046" \
    rename("DialogBox", "ExcelDialogBox") \
    rename("RGB", "ExcelRGB") \
    rename("CopyFile", "ExcelCopyFile") \
    rename("ReplaceText", "ExcelReplaceText") \
    no_auto_exclude
