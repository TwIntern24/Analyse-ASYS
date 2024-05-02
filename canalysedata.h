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

    void on_pbExportExcel_clicked();

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

    void on_pbPrint_clicked();

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

private:
    enum eSHEET
    {
        AnalyseARM = 1,
        AnalyseDM,
        AnalyseZT,
        Label,
        Data
    };
    struct sPROTOCOLITEM
    {
        eSHEET Sheet;
        QVariant Value;
        QString Cell;
        void (CAnalyseData::* p_func)(QAxObject* workbook, sPROTOCOLITEM item); // Member function pointer.
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



    Ui::CAnalyseData *ui;
    QAxObject* m_objExcel = nullptr;
    QAxObject* m_objWorkbook = nullptr;
    QAxObject* m_objExcel1 = nullptr;
    QAxObject* m_objWorkbook1 = nullptr;
    // ARM
    QTextDocument *m_ptdDocumentARM = Q_NULLPTR;

    //QButtonGroup *m_pgbRobotTypeARM;
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
    QString m_filePathExcelTmp = "";
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
    void writeRobotTypeItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeAdviceItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeCauserItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeZTSNFormerItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeRepairOrNotItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeUpgradeItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeCommentsItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeGeneralItem(QAxObject* workbook, sPROTOCOLITEM item);
    void writeRobotType(QAxObject *workbook, sPROTOCOLITEM item);
    void writeCurrentItem(QAxObject *workbook, sPROTOCOLITEM item); //*****************

    void insertImgItem(QAxObject* workbook, sPROTOCOLITEM item);
    void insertRemarkImg(QAxObject* workbook, sPROTOCOLITEM item);  //***************
    void chkTextCokorItem(QAxObject* workbook, sPROTOCOLITEM item);
    void getRowColumn(QString cell, int* row, int* column);
    void buildProtocolTable( void);
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

    //************** NEW ****************************

    struct sREPAIRITEM
    {
        QVariant Value;
        QString Cell;
        void (CAnalyseData::* p_func)( QAxObject* workbook, sREPAIRITEM item );

    };

    QString m_filePathXls = "D:\\Data\\Jana\\Work\\Emily\\Tials\\CreateFiles\\_files\\Repair.xlsx";
    QString m_filePathXlsTemp = "D:\\Data\\Jana\\Work\\Emily\\Tials\\CreateFiles\\_files\\Repair_temp.xls";

    QVector<sREPAIRITEM> vecRepairItems;
    void buildRepairTable( void );

    void writeAmount( QAxObject* workbook, sREPAIRITEM item );
    void writeInforamtion( QAxObject* workbook, sREPAIRITEM item );
    void extendInforamtion( QAxObject* workbook, sREPAIRITEM item );

};

#endif // CANALYSEDATA_H
