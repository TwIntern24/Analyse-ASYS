#ifndef CANALYSEDATA_H
#define CANALYSEDATA_H

#include <QMainWindow>
//#include<QStandardItemModel>
#include <QTextDocument>
#include <QTextBlock>
#include <QButtonGroup>
#include <QSettings>
#include <QObject>
#include <QAxObject>
#include <QVector>
#include <QStandardPaths>
#include <QFileDialog>
#include <QRadioButton>
#include "mainDef.h"
#include "datamatrixgeneratorlib.h"

namespace Ui {
class CAnalyseData;
}

class CAnalyseData : public QMainWindow
{
    Q_OBJECT

public:
    explicit CAnalyseData(QWidget *parent = 0);
    ~CAnalyseData();    

private slots:
    void on_pbRobotSNHide_clicked();

    void on_pbBack_clicked();

    void on_pbSaveIni_clicked();

    void on_pbImageName_clicked();

    void on_leRepairNRARM_textChanged(const QString &arg1);

    void on_leVacuumARM_textChanged(const QString &arg1);

    void on_leFlowARM_textChanged(const QString &arg1);

    void on_leOAG_textChanged(const QString &arg1);

    void on_leUAG_textChanged(const QString &arg1);

    void on_leGeoRz_textChanged(const QString &arg1);

    void on_leGeoRx_textChanged(const QString &arg1);

    void on_leGeoRy_textChanged(const QString &arg1);

    void on_leGeoDelHeight_textChanged(const QString &arg1);

    void on_le180DegVal_2_textChanged(const QString &arg1);

    void on_le270DegVal_2_textChanged(const QString &arg1);

    void on_leVacuumDM_textChanged(const QString &arg1);

    void on_leFlowDM_textChanged(const QString &arg1);

    void on_leRepPosPAR_textChanged(const QString &arg1);

    void on_leRepPosPATH_textChanged(const QString &arg1);

    void createPrintLabel();

    void on_pBtn_ImgRemark_ARM_clicked();

    void on_pBtn_ImgRemark_DM_clicked();

    void on_pBtn_ImgRemark_ZT_clicked();

    void on_leCommutationTH_textChanged(const QString &arg1);

    void on_leCommutationR_textChanged(const QString &arg1);

    void on_leZeroingPosTH_textChanged(const QString &arg1);

    void on_leZeroingPosR_textChanged(const QString &arg1);

    void on_leZUpSCARA_textChanged(const QString &arg1);

    void on_leZDownSCARA_textChanged(const QString &arg1);

    void on_leZUpSCARANT_textChanged(const QString &arg1);

    void on_leZDownSCARANT_textChanged(const QString &arg1);

    void on_leZUpNXT_textChanged(const QString &arg1);

    void on_leZDownNXT_textChanged(const QString &arg1);

    void on_leZUpFA_textChanged(const QString &arg1);

    void on_leZDownFA_textChanged(const QString &arg1);

    void on_pbExport_clicked();         // Export button/function

private:

/* ------------------------------
 * ------ ANALYSE DATA  -------------
 */

    enum eSHEET
    {
        AnalyseARM = 1,
        AnalyseDM,
        AnalyseZT,
        Label,
        Data
    };
    enum eColorChk
    {
        ARMVaccum = 0,
        ARMFlow,
        ARMOAG,
        ARMUAG,
        ARMRz,
        ARMRx,
        ARMRy,
        ARMDeltaH,
        ARMPAR,
        ARMPATH,

        DMVaccum,
        DMFlow,
        DMMHDeg180,
        DMMHDeg270,

        DMComTH,
        DMCOMR,
        DM0PosTH,
        DM0PosR,

        ZSCCurrentUp,
        ZSCCurrentDown,
        ZNTCurrentUp,
        ZNTCurrentDown,
        ZNXTCurrentUp,
        ZNXTCurrentDown,
        ZFACurrentUp,
        ZFACurrentDown
    };

    struct sPROTOCOLITEM
    {
        eSHEET Sheet;
        QVariant Value;
        QString Cell;
        void (CAnalyseData::* p_func)(QAxObject* workbook, sPROTOCOLITEM item); // Member function pointer.
    };

    Ui::CAnalyseData *ui;
    QAxObject* m_objExcel = nullptr;
    QAxObject* m_objWorkbook = nullptr;
    QAxObject* m_objExcel1 = nullptr;
    QAxObject* m_objWorkbook1 = nullptr;
    // ARM
    QTextDocument *m_ptdDocumentARM = Q_NULLPTR;

    //QButtonGroup *m_pgbRobotTypeARM;
    QButtonGroup *m_pgbHDMotorType;
    QButtonGroup *m_pgbAnalyseAdvARM;
    QButtonGroup *m_pgbCauserARM;
    QButtonGroup *m_pgbRepairARMUpgradeChk;
    //QButtonGroup *m_pgbAnalyseChk;
    QButtonGroup *m_pgbSurfaceDamegeChkARM;
    QButtonGroup *m_pgbMagAttChkARM;
    QButtonGroup *m_pgbEleChkARM;
    QButtonGroup *m_pgbGeoChkARM;
    QButtonGroup *m_pgbFunChkARM;
    QButtonGroup *m_pgbDataTransChkARM;

    // DM
    QTextDocument *m_ptdDocumentDM = Q_NULLPTR;

    QButtonGroup *m_pgbAnalyseAdvDM;
    QButtonGroup *m_pgbCauserDM;
    QButtonGroup *m_pgbRepairDMUpgradeChk;
    //QButtonGroup *m_pgbAnalyseChkDM;
    QButtonGroup *m_pgbTHMotorChkDM;
    QButtonGroup *m_pgbRMotorChkDM;
    QButtonGroup *m_pgbTHGearChkDM;
    QButtonGroup *m_pgbRGearChkDM;
    QButtonGroup *m_pgbTiltChkDM;
    QButtonGroup *m_pgbEncJumpTHChkDM;
    QButtonGroup *m_pgbEncJumpRChkDM;
    QButtonGroup *m_pgbEleChkDM;
    QButtonGroup *m_pgbDataTransChkDM;
    QButtonGroup *m_pgbSurfaceDamegeChkDM;
    QButtonGroup *m_pgbCableHolderChkDM;
    QButtonGroup *m_pgbFunChkDM;
    QButtonGroup *m_pgbConductivityChkDM;

    // ZT
    QTextDocument *m_ptdDocumentZT = Q_NULLPTR;

    QButtonGroup *m_pgbAnalyseAdvZT;
    QButtonGroup *m_pgbCauserZT;
    QButtonGroup *m_pgbRepairZTUpgradeChk;
    //QButtonGroup *m_pgbAnalyseChkZT;
    QButtonGroup *m_pgbLVDTZT;
    QButtonGroup *m_pgbRefSenZT;
    QButtonGroup *m_pgbRunningNoiseZT;
    QButtonGroup *m_pgbZMotorZT;
    QButtonGroup *m_pgbCableZT;
    QButtonGroup *m_pgbSurfaceDamegeChkZT;
    QButtonGroup *m_pgbConductivityChkZT;

    //
    QString m_version = "";
    QString m_fileName = "";
    QString m_filePath = "";
//    QString m_filePathTemp = "";
    QString m_settingFilePath = "";
    QString m_filePathExcel = "";
    QString m_filePathExcelRepair = "";
    QString m_filePathExcelMOM = "";
    QString m_filePathExcelTmp = "";
    QString m_filePathExcelRepairTmp = "";
    QString m_filePathExcelMOMTmp = "";
    QString m_comments4Arm = "";
    QString m_comments4DM = "";
    QString m_comments4ZT = "";    
    QString m_imgPath = "";
    QString m_imgDefaultPath = "";

    QString m_RemarkImgPath_ARM = "";   //****
    QString m_RemarkImgPath_DM = "";   //***
    QString m_RemarkImgPath_ZT = "";   //****
    QString m_RemarkImgFileNameARM = "";    //***
    QString m_RemarkImgFileNameDM = "";    //***
    QString m_RemarkImgFileNameZT = "";    //***

    QString m_fileInsteadOf = "";
    QString m_robotType = "";
    QString m_robotSN = "";
    QString m_fileFrom = "";
    QFileInfo m_fileInfoFrom;
    QString m_fileNameFrom = "";
    QString m_filePathFrom = "";
    QString m_strTmpTest = "";
    QString m_strArray[MAX_ARRAY_SIZE];
    QString m_strTiltValRx[MAXPOS];//0:R = -208mm; 1:R = 0; 2:R = 75;3:R = 174;4:R = 203;5:R = 385;
    QString m_strTiltValRy[MAXPOS];
    QString m_strTiltValH4[MAXPOS];
    QStringList m_listFileName;

    bool m_bChkColor[JUDGEITEMS];

    QVector<sPROTOCOLITEM> vecProtocolItems;

    // Main Functions
    void createAnalyseSheet();
    void buildProtocolTable( void);

    void writeRobotTypeItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeAdviceItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeCauserItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeZTSNFormerItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeRepairOrNotItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeUpgradeItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeCommentsItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeOldNew(QAxObject *workbook, sPROTOCOLITEM item);
    void writeGeneralItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeRobotType(QAxObject *workbook, sPROTOCOLITEM item);
    void writeCurrentItem(QAxObject *workbook, sPROTOCOLITEM item); //*****************

    void insertImgItem(QAxObject* workbook, sPROTOCOLITEM item);
    void insertRemarkImg(QAxObject* workbook, sPROTOCOLITEM item);  //***************
    void chkTextCokorItem(QAxObject* workbook, sPROTOCOLITEM item);
    void getRowColumn(QString cell, int* row, int* column);
    void displaySet( bool bSN);
    void closeExcel(void);
    void progressSave(int value);

    void initVal(void);// initial all values
    void createNewIniFile(void);
    void split3PartFromSN(void);

    void addIniNormalGroup(QString strGroupName);
    void addIniTesterGroup(QString strGroupName);
    void addIniRS485Group(QString strGroupName);
    void addIniProtocolVersionGroup(QString strGroupName);

    // get data from GUI setting
    void getDataFromGUIArm( void);
    void getImagePathFromArm( void);
    void getRobotTypeNoSNNCNRFromArm( void);// get robotSN/old12NC/new12NC/repairNR
    void getArmSNFromArm( void);//armSN
    void getGrayBoxDataFromARM( void);// get grayboxSN/firstDeliveryDate/repairNo/lastRepairDate
    void getHDMotorTypeFromARM( void);// get hdMotorType
    void getAdviceCauserFromARM( void);// get analyseAdviceARM/analyseCauserARM
    void getArmOKUpgradeFromARM( void);// get isARMOk/isUpgradeARM
    void getVacFlowValFromARM( void);// get vacARM/flowARM
    void getGeoDataFromARM( void);// get geoRz/geoRx/geoRy/geoDeletaH/isGeoOk
    void getTestsResultsFromARM( void);// get isEletricityOkARM/isDataTransOkARM/isSurOkARM/isFunOkARM/isMagneticOkARM
    void getRepPosPADataFromARM( void);// get repPosPA_R/repPosPA_TH
    void getAnalyseDataFromARM( void);// get analyserARM/analyseDateARM/isAnalyseOkARM
    void getEndDefectTiltDataFromARM( void);// get rMinus208H4Val/rMinus208RxVal/rMinus208RyVal/
                                            //     rZeroH4Val/rZeroRxVal/rZeroRyVal/
                                            //     r75H4Val/r75RxVal/r75RyVal/
                                            //     r174H4Val/r174RxVal/r174RyVal/
                                            //     r203H4Val/r203RxVal/r203RyVal/
                                            //     r385H4Val/r385RxVal/r385RyVal

    void getULValFromARM( void);// get analyseUVal/analyseVVal
    void getTextDocumentFromArm( void);// get analyseCommentsARM
    // for DM
    void getDataFromGUIDM(void);
    void getDMSNFromDM( void);// get DMSN
    void getAdviceCauserFromDM( void);// get analyseAdviceDM/analyseCauserDM
    void getDMOKUpgradeFromDM( void);// get isDMOk/isUpgradeDM
    void getVacFlowValFromDM( void);// get vacDM/flowDM
    void getAngleDataFromDM( void);// get ang180Val/ang270Val
    void getAngleDataUnitMradFromDM( void);// ang180Val/ang270Val unit mrad
    void getTestsResultsFromDM( void);// get isMotorTHOk/isMotorROk/isGearTHOk/isGearROk/isTiltOk/isEncTHOk/isEncROk/isEletricityOkDM/isDataTransOkDM/isSurOkDM/isFunOkDM/isConductivityOkDM
    void getCommutationDataFromDM( void);// get commutationTH/commutationR
    void getZeroingPosDataFromDM( void);// get zeroingPosTH/zeroingPosR
    void getAnalyseDataFromDM( void);// get analyserDM/analyseDateDM/isAnalyseOkDM
    void getTextDocumentFromDM( void);// get analyseCommentsDM
    void getDMDeliveryRepairDate( void);    //get DM first delivery date and the last repair date from GUI //new ****

    // for ZT
    void getDataFromGUIZT(void);
    void getZTSNFromZT( void);// get ZTSN/ZTSN2
    void getAdviceCauserFromZT( void);// get analyseAdviceZT/analyseCauserZT
    void getZTOKUpgradeFromZT( void);// get isZTOk/isUpgradeZT
    void getMeasureDataFromZT( void);// get zUpSCARAVal/zDNSCARAVal/zUpSCARANTVal/zDNSCARANTVal/zUpFAVal/zDNFAVal/zUpNXTVal/zDNNXTVal
    void getTestsResultsFromZT( void);// get isLVDTOk/isRefSnesorOk/isMotorZOk/isRunningNoiseOk/isCableOk/isSurOkZT/isConductivityOkZT
    void getAnalyseDataFromZT( void);// get analyserZT/analyseDateZT/isAnalyseOkZT
    void getTextDocumentFromZT( void);// get analyseCommentsZT
    void getZTDeliveryRepairDate( void);    //get ZT first delivery date and the last repair date from GUI  //new ***


    // get data from Ini file
    // ARM
    void getDataFromIni4Arm( void);
    void getDataFromIni4ArmRobTypeNo( void);
    void getDataFromIni4ArmNCNR( void);
    void getDataFromIni4ArmSN( void);
    void getDataFromIni4ArmGrayBoxData( void);
    void getDataFromIni4HDMotorType( void);
    void getDataFromIni4ArmAdviceCauser( void);
    void getDataFromIni4ArmOKUpgrade( void);
    void getDataFromIni4ArmVacFlowVal( void);
    void getDataFromIni4ArmGeoData( void);
    void getDataFromIni4ArmTestsResults( void);
    void getDataFromIni4ArmRepPosPAData( void);
    void getDataFromIni4ArmAnalyseData( void);
    void getDataFromIni4ArmEndDefectTiltData( void);
    void getDataFromIni4ArmULVal( void);
    void getDataFromIni4ArmComments( void);
    QString calGeoMaxVal(QString  strName);
    //float mmMradCovert(int len, bool bMm2Mrad, float fVal);
    // DM
    void getDataFromIni4DM( void);

    void getDataFromIni4DMSN( void);
    void getDataFromIni4DMAdviceCauser( void);
    void getDataFromIni4DMOKUpgrade( void);
    void getDataFromIni4DMVacFlowVal( void);
    void getDataFromIni4DMAngleData( void);
    void getDataFromIni4DMTestsResults( void);
    void getDataFromIni4DMCommutationData( void);
    void getDataFromIni4DMZeroingPosData( void);
    void getDataFromIni4DMAnalyseData( void);
    void getDataFromIni4DMComments( void);
    void getDataFromIni4DMDeliveryRepairDate( void);    //get DM first delivery date and the last repair date from ini  //new ***

    //ZT
    void getDataFromIni4ZT( void);
    void getDataFromIni4ZTSN( void);
    void getDataFromIni4ZTAdviceCauser( void);
    void getDataFromIni4ZTOKUpgrade( void);
    void getDataFromIni4ZTMeasureData( void);
    void getDataFromIni4ZTTestsResults( void);
    void getDataFromIni4ZTAnalyseData( void);

    void getDataFromIni4ZTComments( void);
    void getDataFromIni4ZTDeliveryRepairDate( void);    //get ZT first delivery date and the last repair date from ini  //new ***

    void setRadioButtonsIDsInGB4ARM( void);
    void setRadioButtonsIDsInGB4DM( void);
    void setRadioButtonsIDsInGB4ZT( void);

    DataMatrixGeneratorLib libDataMatrix;
    void insertDataMatrix(QAxObject* workbook, int iWorkSheet, QString strCell, QString strImgFileName, double dImgSize);
    bool isDigitStr(QString strSrc);

    void getLabelData();
    QVector<QString> vecLabelData;

    QString mFilePathLabel = "";
    QAxObject* m_objLabelExcel = nullptr;
    QAxObject* m_objLabelWorkbook = nullptr;
    QString mFilePathLabelTemplate = "";

    void createLabelFile();

    //*******************************************************
    //************** REPAIR DATA ****************************

    struct sREPAIRITEM  // struct for Repair Matrix items
    {
        QVariant Value;
        QString Cell;
        void (CAnalyseData::* p_func)( QAxObject* workbook, sREPAIRITEM item );
    };

    // VARIABLES ------------------------
    QVector<sREPAIRITEM> vecRepairItems;    // Vector for saving Repair Matrix items
    // arm
    QButtonGroup *m_pgbRepairARM_ArmBelts;
    QButtonGroup *m_pgbRepairARM_UpperArmHousingUpgrade;
    QButtonGroup *m_pgbRepairARM_UpperArmHousing;
    QButtonGroup *m_pgbRepairARM_UpperArmLid;
    QButtonGroup *m_pgbRepairARM_LowerArmHousingUpgrade;
    QButtonGroup *m_pgbRepairARM_LowerArmHousing;
    QButtonGroup *m_pgbRepairARM_LowerArmLid;
    QButtonGroup *m_pgbRepairARM_ArmDriveInterface;
    QButtonGroup *m_pgbRepairARM_ArmGripperInterfaceScara;
    QButtonGroup *m_pgbRepairARM_ArmGripperInterfaceNT;
    QButtonGroup *m_pgbRepairARM_BeltReel;
    QButtonGroup *m_pgbRepairARM_TorxScrew;
    QButtonGroup *m_pgbRepairARM_Bearings;
    QButtonGroup *m_pgbRepairARM_RepairIn;
    // DM
    QButtonGroup *m_pgbRepairDM_DMLikaMotor;
    QButtonGroup *m_pgbRepairDM_CableHood;
    QButtonGroup *m_pgbRepairDM_DMHousing;
    QButtonGroup *m_pgbRepairDM_DMLid;
    QButtonGroup *m_pgbRepairDM_SlipRing;
    QButtonGroup *m_pgbRepairDM_HollowShaft;
    QButtonGroup *m_pgbRepairDM_RepairIn;
    // ZT
    QButtonGroup *m_pgbRepairZT_ZStroke35;
    QButtonGroup *m_pgbRepairZT_ZStroke50;
    QButtonGroup *m_pgbRepairZT_ZMHousingScara;
    QButtonGroup *m_pgbRepairZT_ZMHousingNT;
    QButtonGroup *m_pgbRepairZT_GuidingShaftsScara;
    QButtonGroup *m_pgbRepairZT_GuidingShaftsNT;
    QButtonGroup *m_pgbRepairZT_SmallGuidingShafts;
    QButtonGroup *m_pgbRepairZT_ClampingFlange;
    QButtonGroup *m_pgbRepairZT_AdapterCable;
    QButtonGroup *m_pgbRepairZT_RepairIn;

    // FUNCTIONS ---------------------------
    // create buttons groups for repair data
    void setRadioButtonsIDsInGB4RepairARM( void);
    void setRadioButtonsIDsInGB4RepairDM( void);
    void setRadioButtonsIDsInGB4RepairZT( void);

    // Main Functions
    void createRepairMatrix();
    void buildRepairTable( void );

    // functions for saving repair data in .ini file
    void getDataFromRepair( void );
    void getRepairFromARM( void );
    void getRepairFromDM( void );
    void getRepairFromZT( void );

    // functions for loading repair data from .ini file
    void getDataFromIni4Repair( void);
    void getDataFromIni4ARMRepair( void);
    void getDataFromIni4DMRepair( void);
    void getDataFromIni4ZTRepair( void);

    // helper functions for writing in Excel
    void writeAmount( QAxObject* workbook, sREPAIRITEM item );
    void writeInformation( QAxObject* workbook, sREPAIRITEM item );
    void extendInformation( QAxObject* workbook, sREPAIRITEM item );
    void writeRepairIn(QAxObject *workbook, sREPAIRITEM item);


    //*******************************************************
    //************** MOM SHEET ******************************

    // VARIABLES ---------------------------
    QString strResAnalysis;     // String for text in Results Analysis column in MOM Sheet
    int startPos = 0;       // variable for char counting

    // FUNCTIONS ---------------------------
    // Main Function
    void createMOMSheet( void );

    // write MOM Sheet
    void writeGeneralInfoMOM( void );
    void writeArmInfoMOM( void );
    void writeDMInfoMOM( void );
    void writeZInfoMOM( void );

    // helper functions for char counting in string
    void getNextLineIdx ( void );
    void insertAndReturnLastIdx(QString insertText);
    void removeAndReturnLastIdx(int numberRemove);

};

#endif // CANALYSEDATA_H
