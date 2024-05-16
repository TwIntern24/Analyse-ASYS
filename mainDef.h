#ifndef MAINDEF_H
#define MAINDEF_H

#define USEAPPLICATIONPATH

#define BEGIN_ARMINFO1 "ARM1_Info"
#define BEGIN_DMINFO1 "DM1_Info"
#define BEGIN_ZTINFO "ZT_Info"

#define BEGIN_ROBOTINFO "Robot_Info"
#define BEGIN_CHECKSINFO "Checks_Info"
#define BEGIN_DMINFO "DM_Info"
#define BEGIN_ZAXEINFO "ZAxe_Info"
#define BEGIN_PACKINFO "Pack_Info"
#define BEGIN_ARMINFO "Arm_Info"
#define BEGIN_TESTERINFO "Robot_Testers"
#define BEGIN_RS485TEST "RS485_Test"
#define BEGIN_PROTOVOLVERSION "ProtocolVersion"
#define BEGIN_ARMREPAIR "Arm_Repair"
#define BEGIN_DMREPAIR "DM_Repair"
#define BEGIN_ZTREPAIR "ZT_Repair"
#define NG_VALUE "9999"
#define JUDGEITEMS 26

#define ImageResizeWidth 265
#define ImageResizeHeight 202

#define ARMVACUUM_LIMIT 0.5
#define ARMFLOW_UP 8
#define ARMFLOW_DN 4.2
#define OAG_UP 435
#define OAG_DN 415
#define UAG_UP 365
#define UAG_DN 345
#define Rz_UP 0.5
#define Rz_DN -0.5
#define Rx_UP 0.7
#define Rx_DN -0.7
#define Ry_UP 0.6
#define Ry_DN -0.6
#define DeltaHeight_UP 0.13
#define DeltaHeight_DN -0.14
#define REPPOSPA_UP 600
#define REPPOSPA_DN -600
#define MicroHite_UP 0.05
#define MicroHite_DN -0.05
#define TINYVALUE 0.0001
#define FLOAT_PRECISON 3

#define TOOL_LENGTH 525
#define TOOL_LENGTH1 100

#define ROBOTTYPE_DF "DF"
#define ROBOTTYPE_NT "AAR-NT"
#define ROBOTTYPE_NXT "AAR-NXT"
#define ROBOTTYPE_SC "AAR"
//#define ROBOTTYPE_NT "NT"
//#define ROBOTTYPE_NXT "NXT"
//#define ROBOTTYPE_SC "SC"

#define MAX_ARRAY_SIZE 151
#define MAXPOS 6

enum ERobotTypes
{
       Robot_AAR,
       Robot_AARNT,
       Robot_NXT
};

enum ETestResult
{
       Test_OK,
       Test_NG,
       Test_NA      //new code******************************
};

enum EOldNew
{
    OldNew_NEW,
    OldNew_OLD,
    OldNew_NA=-1
};

enum EHDMotorType
{
    V1,
    V0,
    DFV1
};

enum EAnalyseAdvice
{
       NFF,
       Warrenty,
       GoodWill,
       WithCosts,
       ScrapItem
};

enum ECauserChk
{
       Customer,
       ASYS
};

enum EZTLength
{
       Len35mm,
       Len50mm,
       Len65mm
};

enum ERepair
{
    Repair_NA,
    Repair_OK,
    Repair_Repair
};

enum ERepairIn
{
    RepairIn_TW,
    RepairIn_EU
};

#endif //MAINDEF_H
