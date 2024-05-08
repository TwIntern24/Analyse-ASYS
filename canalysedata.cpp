#include "canalysedata.h"
#include "ui_canalysedata.h"

#include <QMessageBox>
#include <QDebug>
#include <QDateTime>
#include <QtMath>
#include <QSettings>
#include <QDir>
#include <QStringList>

#include <QPrinter>
#include <QPrintDialog>
#include <QPrinterInfo>
#include <QDoubleValidator>

CAnalyseData::CAnalyseData(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::CAnalyseData)
{
    ui->setupUi(this);
    m_version = "2.0.3.1";
    this->setWindowTitle( "Analyse-ASYS Ver. "+ m_version);

    displaySet( true);
    setRadioButtonsIDsInGB4ARM();
    setRadioButtonsIDsInGB4DM();
    setRadioButtonsIDsInGB4ZT();
    setRadioButtonsIDsInGB4RepairARM();
    setRadioButtonsIDsInGB4RepairDM();
    setRadioButtonsIDsInGB4RepairZT();

    ui->gbTiltDataChk->setVisible(false);
    ui->leRobotSN1st->setFocus();   
    ui->stackedWidget->setCurrentIndex(0);
    ui->stackedWidget_Image->setCurrentIndex(0);

    ui->pbBack->setStyleSheet("font-size: 16px;font-weight: bold;");
    ui->pbExport->setStyleSheet("font-size: 16px;font-weight: bold;");
    ui->pbSaveIni->setStyleSheet("font-size: 16px;font-weight: bold;");




}

CAnalyseData::~CAnalyseData()
{
    closeExcel();
    delete ui;
}

void CAnalyseData::initVal(void)
{
   //-- clear all value --//
    ui->leRobotSN->clear();
    ui->leARMSN->clear();
    ui->leOld12NC->clear();
    ui->leNew12NC->clear();
    ui->leRepairNRARM->clear();
    ui->leVacuumARM->clear();
    ui->leFlowARM->clear();
    ui->leGeoRz->clear();
    ui->leGeoRx->clear();
    ui->leGeoRy->clear();
    ui->leOAG->clear();
    ui->leUAG->clear();
    ui->lbImagePath->clear();
    ui->leBIDCode->clear();// 20230706 add
    ui->lbImageName->clear();
    ui->leGeoDelHeight->clear();
    ui->leRepPosPAR->clear();
    ui->leRepPosPATH->clear();
    ui->leAnalysePerformerARM->clear();
    ui->leAnalyseDateARM->clear();
    ui->leGrayboxSN->clear();
    ui->leFirstDeliveryDate->clear();
    ui->leRepairNo->clear();
    ui->leLastRepairDate->clear();
    ui->teARMComments->clear();
    ui->leDMSN->clear();
    ui->leRepairNRARM->clear();
    ui->leVacuumDM->clear();
    ui->leFlowDM->clear();
    ui->le180DegVal->clear();
    ui->le270DegVal->clear();
    ui->le180DegVal_2->clear();
    ui->le270DegVal_2->clear();
    ui->leCommutationTH->clear();
    ui->leCommutationR->clear();
    ui->leZeroingPosTH->clear();
    ui->leZeroingPosR->clear();
    ui->leAnalysePerformerDM->clear();
    ui->leAnalyseDateDM->clear();
    ui->teDMComments->clear();
    ui->leZTSN2->clear();
    ui->leRepairNRARM->clear();
    ui->leZUpSCARA->clear();
    ui->leZDownSCARA->clear();
    ui->leZUpSCARANT->clear();
    ui->leZDownSCARANT->clear();
    ui->leZUpFA->clear();
    ui->leZDownFA->clear();
    ui->leZUpNXT->clear();
    ui->leZDownNXT->clear();
    ui->leAnalysePerformerZT->clear();
    ui->leAnalyseDateZT->clear();
    ui->teZTComments->clear();
    ui->cbUpgradeARM->setChecked(false);
    ui->cbUpgradeDM->setChecked(false);
    ui->cbUpgradeZT->setChecked(false);
    ui->rbNFFARM->setChecked(true);
    ui->rbCustomerARM->setChecked( true);
    ui->rbRepairARM->setChecked( true);
    ui->rbSurfNotDamegeNOKARM->setChecked( true);
    ui->rbMagAttNOK->setChecked( true);
    ui->rbEletricityNOKARM->setChecked( true);
    ui->rbGeoNOK->setChecked( true);
    ui->rbFunTestNOKARM->setChecked( true);
    ui->rbDataTransNOKARM->setChecked( true);
    ui->rbNFFDM->setChecked(true);
    ui->rbCustomerDM->setChecked( true);
    ui->rbRepairDM->setChecked( true);
    ui->rbSurfNotDamegeNOKDM->setChecked( true);
    ui->rbEletricityNOKDM->setChecked( true);
    ui->rbFunTestNOKDM->setChecked( true);
    ui->rbDataTransTestNOKDM->setChecked( true);
    ui->rbTHMotorNOK->setChecked( true);
    ui->rbRMotorNOK->setChecked( true);
    ui->rbTHGearNOK->setChecked( true);
    ui->rbRGearNOK->setChecked( true);
    ui->rbTiltNOK->setChecked( true);
    ui->rbEncJumpTestTHNOK->setChecked( true);
    ui->rbEncJumpTestRNOK->setChecked( true);
    ui->rbConductivityChkNOKDM->setChecked( true);
    ui->rbNFFZT->setChecked(true);
    ui->rbCustomerZT->setChecked( true);
    ui->rbRepairZT->setChecked( true);
    ui->rbSurfNotDamegeNOKZT->setChecked( true);
    ui->rbLVDTNOK->setChecked( true);
    ui->rbRefSensorNOK->setChecked( true);
    ui->rbZMotorNOK->setChecked( true);
    ui->rbRunNoiseNOK->setChecked( true);
    ui->rbCableNOK->setChecked( true);
    ui->rbConductivityChkNOKZT->setChecked( true);
    ui->leArmFirstDelivery->clear();
    ui->leArmLastRepair->clear();
    ui->leDMFirstDelivery->clear();
    ui->leDMLastRepair->clear();
    ui->leZTFirstDelivery->clear();
    ui->leZTLastRepair->clear();

    ui->rbArmBeltsAvrPrice_NA->setChecked(true);
    ui->rbUpperArmHousingUpgrade_NA->setChecked(true);
    ui->rbUpperArmHousing_NA->setChecked(true);
    ui->rbUpperArmLid_NA->setChecked(true);
    ui->rbLowerArmHousingUpgrade_NA->setChecked(true);
    ui->rbLowerArmHousing_NA->setChecked(true);
    ui->rbLowerArmLid_NA->setChecked(true);
    ui->rbArmDriveInterface_NA->setChecked(true);
    ui->rbArmGripperInterfaceScara_NA->setChecked(true);
    ui->rbArmGripperInterfaceNT_NA->setChecked(true);
    ui->rbBeltReel_NA->setChecked(true);
    ui->rbTorxScrew_NA->setChecked(true);
    ui->rbBearings_NA->setChecked(true);
    ui->cbArmTW->setChecked(true);
    ui->rbDMLikaMotor_NA->setChecked(true);
    ui->rbCableHood_NA->setChecked(true);
    ui->rbDMHousing_NA->setChecked(true);
    ui->rbDMLid_NA->setChecked(true);
    ui->rbSlipRing_NA->setChecked(true);
    ui->rbHollowShaft_NA->setChecked(true);
    ui->cbDMTW->setChecked(true);
    ui->rbZStroke35_NA->setChecked(true);
    ui->rbZStroke50_NA->setChecked(true);
    ui->rbZMHousingScara_NA->setChecked(true);
    ui->rbZMHousingNT_NA->setChecked(true);
    ui->rbGuidingShaftsScara_NA->setChecked(true);
    ui->rbGuidingShaftsNT_NA->setChecked(true);
    ui->rbSmallGuidingShafts_NA->setChecked(true);
    ui->rbClampingFlange_NA->setChecked(true);
    ui->rbAdapterCable_NA->setChecked(true);
    ui->cbZTTW->setChecked(true);
}

//
void CAnalyseData::addIniTesterGroup(QString strGroupName)
{
    QSettings settings(m_filePath, QSettings::IniFormat);
    QString strArray1[MAX_ARRAY_SIZE];
    QString strArray2[MAX_ARRAY_SIZE];
    QString strVal = "";

    settings.beginGroup(strGroupName);

    for(int i = 0; i < MAX_ARRAY_SIZE; i++)
    {
        if(i==MAX_ARRAY_SIZE-1)
        {
            m_strArray[i] = QString("size");
            strVal = "150";
            settings.setValue(m_strArray[i], strVal);
        }
        else
        {
            m_strArray[i] = QString("%1/Tester").arg(i+1);
            strArray1[i] = QString("%1/Notes").arg(i+1);
            strArray2[i] = QString("%1/Time").arg(i+1);
            strVal = "";

            settings.setValue(m_strArray[i], strVal);
            settings.setValue(strArray1[i], strVal);
            settings.setValue(strArray2[i], strVal);
        }
    }
    settings.endGroup();
}

void CAnalyseData::addIniRS485Group(QString strGroupName)
{
    QSettings settings(m_filePath, QSettings::IniFormat);
    QString str1 = "size";
    QString strVal = "0";

    settings.beginGroup(strGroupName);
    settings.setValue(str1, strVal);
    settings.endGroup();
}

void CAnalyseData::addIniProtocolVersionGroup(QString strGroupName)
{
    QSettings settings(m_filePath, QSettings::IniFormat);
    QString str1 = "Version";
    QString str2 = "AppVersion";
    QString str3 = "Last_saved";
    QString strVal = "";
    settings.beginGroup(strGroupName);
    settings.setValue(str1, strVal);
    settings.setValue(str2, strVal);
    settings.setValue(str3, strVal);
    settings.endGroup();
}

void CAnalyseData::addIniNormalGroup(QString strGroupName)
{
    QSettings settings(m_filePath, QSettings::IniFormat);
    QString strVal = "";
    settings.beginGroup(strGroupName);
    //qDebug()<< "strGroupName = "<<strGroupName<<endl;
    for(int i = 0; i < MAX_ARRAY_SIZE; i++)
    {
        if(i==MAX_ARRAY_SIZE-1)
        {
            m_strArray[i] = QString("size");
            strVal = "150";
        }
        else
        {
            m_strArray[i] = QString("%1/val").arg(i+1);
            if(strGroupName==BEGIN_ROBOTINFO)
            {
                if(i>=1&&i<=7)// 2\val~8\val
                {
                    strVal = "ANALYSE";
                }
                else if(i==14||i==15)// 15\val,16\val
                {
                    strVal = "1";
                }
                else
                {
                    strVal = "";
                }
            }
            else if(strGroupName==BEGIN_CHECKSINFO)
            {
                if((i>=0&&i<=5)||i==15||i==16||(i>=18&&i<=20))// 1\val~6\val, 16\val~17\val, 19\val~21\val
                {
                    strVal = "true";
                }
//                else if(i==15||i==16)// 16\val~17\val
//                {
//                    strVal = "true";
//                }
                else
                {
                    strVal = "false";
                }
            }
            else
            {
                strVal = "9999";
            }
        }
        settings.setValue(m_strArray[i], strVal);
    }
    settings.endGroup();
}

void CAnalyseData::createNewIniFile(void)
{
    QString strGroup[6] = {BEGIN_ROBOTINFO, BEGIN_CHECKSINFO, BEGIN_DMINFO, BEGIN_ZAXEINFO, BEGIN_PACKINFO, BEGIN_ARMINFO};

    for(int j = 0; j < 6; j++)
    {
        addIniNormalGroup(strGroup[j]);
    }
    addIniTesterGroup(BEGIN_TESTERINFO);
    addIniRS485Group(BEGIN_RS485TEST);
    addIniProtocolVersionGroup(BEGIN_PROTOVOLVERSION);
}

void CAnalyseData::split3PartFromSN(void)
{ // ANALYSE_AAR-NT2468_KW2316 >>> [0]:ANALYSE->Protocol; [1]:AAR-NT2468; [2]:KW2316
    bool bRst = false;
    QString strName = "1/val";
    int iLen = 0;
    int ilenSN = 0;
    m_fileInsteadOf = "Protocol";
    bRst = m_fileName.contains("_",Qt::CaseSensitive);
    if( bRst==true)
    {
      m_listFileName = m_fileName.split("_");
      //qDebug()<<"listFileName-0:"<<m_listFileName[0]<<"listFileName-1:"<<m_listFileName[1]<<"listFileName[2]:"<<m_listFileName[2]<<endl;
      m_fileInsteadOf += "_" + m_listFileName[1] + "_" + m_listFileName[2];// Protocol_AAR-NT2468_KW2316>>write to [Robot_Info]>>1\val
      QSettings settings(m_filePath, QSettings::IniFormat);
      settings.beginGroup(BEGIN_ROBOTINFO);
      settings.setValue(strName, m_fileInsteadOf);
      QString strType_SN = m_listFileName[1];
      QString strSNValid = "";
      for (QChar ch : strType_SN)
      {
        if (ch.isDigit())
            strSNValid.append(ch);
      }
      ilenSN = strSNValid.size();
      iLen = m_listFileName[1].length();
      m_robotType = m_listFileName[1].left(iLen-ilenSN);

      while(strSNValid.size() < 4)
      {
          strSNValid.insert(0, "0");
      }
      m_robotSN = strSNValid;   //m_robotSN = m_listFileName[1].right(ilenSN);  //robot serial number
      ui->lbRobotType->setText(m_robotType);
      ui->leRobotTypeSN->setText(m_robotSN);
    }
    else
    {
      QMessageBox::critical(NULL, "Error", "File Name not correct", QMessageBox::Yes, QMessageBox::Yes);
      return;
    }
}

void CAnalyseData::buildProtocolTable( void)
{
                               // Sheet                  // Value                                                 // Cell             // Function
    //---- ARM -----------------------------------------------------------------------------------------------------------------------------------------//
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->lbRobotType->text(),                                       "B1",           nullptr});  //Changed ***
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leRobotTypeSN->text(),                                     "C1",           nullptr});  //Changed ***
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leArmFirstDelivery->text(),                                "E3",           nullptr});  //new ***
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leArmLastRepair->text(),                                   "E4",           nullptr});  //new ***
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leRobotSN->text(),                                         "D7",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leARMSN->text(),                                           "D8",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leOld12NC->text(),                                         "D10",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leNew12NC->text(),                                         "D11",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leRepairNRARM->text(),                                     "D13",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leVacuumARM->text(),                                       "D14",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leFlowARM->text(),                                         "D15",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leOAG->text(),                                             "D16",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leUAG->text(),                                             "D17",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leGeoRz->text(),                                           "D18",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leGeoRx->text(),                                           "D19",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leGeoRy->text(),                                           "D20",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leGeoDelHeight->text(),                                    "D21",          chkTextCokorItem});
//    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbGeoChkARM->checkedId(),                                   "E22",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbGeoChkARM->checkedId(),                                   "D22",          nullptr});  //Changed ***
//    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbEleChkARM->checkedId(),                                   "E23",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbEleChkARM->checkedId(),                                   "D23",          nullptr});  //Changed ***
//    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbDataTransChkARM->checkedId(),                             "E24",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbSurfaceDamegeChkARM->checkedId(),                         "E25",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbFunChkARM->checkedId(),                                   "E26",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbDataTransChkARM->checkedId(),                             "D24",          nullptr});  //****
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbSurfaceDamegeChkARM->checkedId(),                         "D25",          nullptr});  //****
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbFunChkARM->checkedId(),                                   "D26",          nullptr});  //****
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leRepPosPAR->text(),                                       "D27",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leRepPosPATH->text(),                                      "D28",          chkTextCokorItem});
//    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbMagAttChkARM->checkedId(),                                "E29",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbMagAttChkARM->checkedId(),                                "D29",          nullptr});  //****
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leAnalysePerformerARM->text(),                             "D33",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leAnalyseDateARM->text(),                                  "D34",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leGrayboxSN->text(),                                       "I1",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leFirstDeliveryDate->text(),                               "I2",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leRepairNo->text(),                                        "I3",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->leLastRepairDate->text(),                                  "I4",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbAnalyseAdvARM->checkedId(),                               "F7",           writeAdviceItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbCauserARM->checkedId(),                                   "H7",           writeCauserItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_pgbRepairARMUpgradeChk->checkedId(),                         "I10",          writeRepairOrNotItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     ui->cbUpgradeARM->isChecked(),                                 "I11",          writeUpgradeItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_comments4Arm,                                                "H14",          writeCommentsItem});
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_imgPath,                                                     "B36",          insertImgItem});//B45
    vecProtocolItems.append({eSHEET::AnalyseARM,     m_RemarkImgPath_ARM,                                           "G36",          insertRemarkImg});//new code ******

    //---- DM -----------------------------------------------------------------------------------------------------------------------------------------//
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->lbRobotType->text(),                                       "B1",           nullptr});  //Changed****
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leRobotTypeSN->text(),                                     "C1",           nullptr});  //Changed***
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leDMFirstDelivery->text(),                                 "E3",           nullptr});  //new ***
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leDMLastRepair->text(),                                    "E4",           nullptr});  //new ***
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leRobotSN->text(),                                         "D7",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leDMSN->text(),                                            "D9",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leRepairNRARM->text(),                                     "D13",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leVacuumDM->text(),                                        "D14",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leFlowDM->text(),                                          "D15",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->le180DegVal->text(),                                       "D17",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->le270DegVal->text(),                                       "D18",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->le180DegVal_2->text(),                                     "F17",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->le270DegVal_2->text(),                                     "F18",          chkTextCokorItem});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbTHMotorChkDM->checkedId(),                                "E19",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbRMotorChkDM->checkedId(),                                 "E20",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbTHGearChkDM->checkedId(),                                 "E21",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbRGearChkDM->checkedId(),                                  "E22",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbTiltChkDM->checkedId(),                                   "E23",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbTHMotorChkDM->checkedId(),                                "D19",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbRMotorChkDM->checkedId(),                                 "D20",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbTHGearChkDM->checkedId(),                                 "D21",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbRGearChkDM->checkedId(),                                  "D22",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbTiltChkDM->checkedId(),                                   "D23",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leCommutationTH->text(),                                   "D24",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leCommutationR->text(),                                    "D25",          chkTextCokorItem});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbEncJumpTHChkDM->checkedId(),                              "E26",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbEncJumpRChkDM->checkedId(),                               "E27",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbEleChkDM->checkedId(),                                    "E28",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbDataTransChkDM->checkedId(),                              "E29",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbSurfaceDamegeChkDM->checkedId(),                          "E30",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbFunChkDM->checkedId(),                                    "E31",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbEncJumpTHChkDM->checkedId(),                              "D26",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbEncJumpRChkDM->checkedId(),                               "D27",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbEleChkDM->checkedId(),                                    "D28",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbDataTransChkDM->checkedId(),                              "D29",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbSurfaceDamegeChkDM->checkedId(),                          "D30",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbFunChkDM->checkedId(),                                    "D31",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leZeroingPosTH->text(),                                    "D32",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leZeroingPosR->text(),                                     "D33",          chkTextCokorItem});
//    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbConductivityChkDM->checkedId(),                           "E34",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbConductivityChkDM->checkedId(),                           "D34",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leAnalysePerformerDM->text(),                              "D38",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->leAnalyseDateDM->text(),                                   "D39",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbAnalyseAdvDM->checkedId(),                                "F7",           writeAdviceItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbCauserDM->checkedId(),                                    "H7",           writeCauserItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_pgbRepairDMUpgradeChk->checkedId(),                          "I10",          writeRepairOrNotItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      ui->cbUpgradeDM->isChecked(),                                  "I11",          writeUpgradeItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_comments4DM,                                                 "H14",          writeCommentsItem});
    vecProtocolItems.append({eSHEET::AnalyseDM,      m_RemarkImgPath_DM,                                            "G36",          insertRemarkImg});//new code ******

    //---- ZT -----------------------------------------------------------------------------------------------------------------------------------------//
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->lbRobotType->text(),                                       "B1",           nullptr});  //Changed ***
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leRobotTypeSN->text(),                                     "C1",           nullptr});  //Chenged ***
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZTFirstDelivery->text(),                                 "E3",           nullptr});  //new ***
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZTLastRepair->text(),                                    "E4",           nullptr});  //new ***
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leRobotSN->text(),                                         "D7",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZTLen->text()+"-XX-XX-",                                 "C9",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZTSN2->text(),                                           "D9",           nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leRepairNRARM->text(),                                     "D13",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZUpSCARA->text(),                                        "F14",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZDownSCARA->text(),                                      "F15",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZUpSCARANT->text(),                                      "F16",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZDownSCARANT->text(),                                    "F17",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZUpNXT->text(),                                          "F18",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZDownNXT->text(),                                        "F19",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZUpFA->text(),                                           "F20",          chkTextCokorItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leZDownFA->text(),                                         "F21",          chkTextCokorItem});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbLVDTZT->checkedId(),                                      "E22",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbRefSenZT->checkedId(),                                    "E23",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbZMotorZT->checkedId(),                                    "E24",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbRunningNoiseZT->checkedId(),                              "E25",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbCableZT->checkedId(),                                     "E26",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbSurfaceDamegeChkZT->checkedId(),                          "E27",          nullptr});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbConductivityChkZT->checkedId(),                           "E28",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbLVDTZT->checkedId(),                                      "D22",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbRefSenZT->checkedId(),                                    "D23",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbZMotorZT->checkedId(),                                    "D24",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbRunningNoiseZT->checkedId(),                              "D25",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbCableZT->checkedId(),                                     "D26",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbSurfaceDamegeChkZT->checkedId(),                          "D27",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbConductivityChkZT->checkedId(),                           "D28",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leAnalysePerformerZT->text(),                              "E32",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->leAnalyseDateZT->text(),                                   "E33",          nullptr});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbAnalyseAdvZT->checkedId(),                                "F7",           writeAdviceItem});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbCauserZT->checkedId(),                                    "H7",           writeCauserItem}); *******
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbCauserZT->checkedId(),                                    "I7",           writeCauserItem});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbRepairZTUpgradeChk->checkedId(),                          "I10",          writeRepairOrNotItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_pgbRepairZTUpgradeChk->checkedId(),                          "J10",          writeRepairOrNotItem}); //****
//    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->cbUpgradeZT->isChecked(),                                  "I11",          writeUpgradeItem});
    vecProtocolItems.append({eSHEET::AnalyseZT,      ui->cbUpgradeZT->isChecked(),                                  "J11",          writeUpgradeItem}); //*****
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_comments4ZT,                                                 "I14",          writeCommentsItem});
//    vecProtocolItems.append({eSHEET::AnalyseZT,      m_RemarkImgPath_ZT,                                            "H36",          insertRemarkImg});
    vecProtocolItems.append({eSHEET::AnalyseZT,      m_RemarkImgPath_ZT,                                            "I36",          insertRemarkImg});  //new code ******

    //---- Data -----------------------------------------------------------------------------------------------------------------------------------------//
    //m_robotType
    vecProtocolItems.append({eSHEET::Data,           ui->lbRobotType->text(),                                        "F1",          writeRobotType});  //***********
    vecProtocolItems.append({eSHEET::Data,           ui->lbRobotType->text()+ui->leRobotTypeSN->text(),              "A4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leRMinus208H4->text(),                                      "B4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leRMinus208Rx->text(),                                      "C4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leRMinus208Ry->text(),                                      "D4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR0H4->text(),                                             "E4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR0Rx->text(),                                             "F4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR0Ry->text(),                                             "G4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR75H4->text(),                                            "H4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR75Rx->text(),                                            "I4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR75Ry->text(),                                            "J4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR174H4->text(),                                           "K4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR174Rx->text(),                                           "L4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR174Ry->text(),                                           "M4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR203H4->text(),                                           "N4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR203Rx->text(),                                           "O4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR203Ry->text(),                                           "P4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR385H4->text(),                                           "Q4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR385Rx->text(),                                           "R4",          nullptr});
    vecProtocolItems.append({eSHEET::Data,           ui->leR385Ry->text(),                                           "S4",          nullptr});
}

void CAnalyseData::displaySet( bool bSN)
{
    if( bSN==true)
    {
        ui->gbRobotSN->show();
        ui->stackedWidget->setCurrentIndex(0);  //ui->pbRobotSNHide->show();
        ui->tabWidget->hide();
        ui->pbBack->hide();
        ui->pbSaveIni->hide();
//        ui->pbExportExcel->hide();
        ui->pgbProcess->hide();
        ui->lbStatus->hide();
//        ui->pbPrint->hide();
    }// show the Robot SN page
    else
    {
        ui->gbRobotSN->hide();
        ui->stackedWidget->setCurrentIndex(1);  //ui->pbRobotSNHide->hide();
        ui->tabWidget->show();
        ui->pbBack->show();
        ui->pbSaveIni->show();
//        ui->pbExportExcel->show();
        ui->pgbProcess->show();
        ui->lbStatus->show();
        ui->lbStatus->setText("Status");
//        ui->pbPrint->show();
        ui->tabWidget->setCurrentIndex(0);
    }
}

//---------------------------------------- Next Button clicked ------------------------------//
void CAnalyseData::on_pbRobotSNHide_clicked()
{
    //-------- User input error check --------//
    if( ui->leRobotSN1st->text().isEmpty() || ui->leRobotSN1st->text().size() < 9)  //ANALYSE_
    {
        QMessageBox::critical( NULL, "Error", "Field Empty! Please enter serial number", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }

    if((ui->leRobotSN1st->text().size() >= 9) && (ui->leRobotSN1st->text().size() < 21))   //Not complete
    {
        QStringList strlstSerialNum = ui->leRobotSN1st->text().split("_");  //ANALYSE_AAR-NT2565_KW2322
        if(strlstSerialNum.size() < 3)
        {
            QMessageBox::critical( NULL, "Error", "Serial Number Incomplete!", QMessageBox::Yes, QMessageBox::Yes);
            return;
        }else
        {
            if((strlstSerialNum[0] == "") || (strlstSerialNum[1] == "") || (strlstSerialNum[2] == ""))
            {
                QMessageBox::critical( NULL, "Error", "Serial Number Incomplete!", QMessageBox::Yes, QMessageBox::Yes);
                return;
            }
        }
    }
    //else
    //{
      displaySet( false);
      // Create a QSettings object with the path to the INI file
      m_fileName = ui->leRobotSN1st->text();

      #ifdef USEAPPLICATIONPATH
        m_settingFilePath = QApplication::applicationDirPath()+"\\Setting_debug.ini";    //QApplication::applicationDirPath()
      #else
//        m_settingFilePath = "D:\\ASYS\\Projects\\Analyse_ASYS\\Setting.ini";
          m_settingFilePath = "D:\\Data\\twintern\\Jana\\Work\\Emily\\Analysis_ASYS_Material\\Analyse_ASYS_ver2.0.3.1\\Analyse_ASYS\\release"
      #endif
      QSettings settings(m_settingFilePath,QSettings::IniFormat);
      settings.beginGroup("PathSetting");

      QString strExcelTemp = settings.value("TemplateVersion").toString();  //ANALYSETILT_v1.0  //************
//      qDebug() << strExcelTemp;
      m_filePathExcelTmp = QDir::toNativeSeparators(QApplication::applicationDirPath()) + "\\" + strExcelTemp + ".xls";     //*************
//      qDebug() << m_filePathExcelTmp;

      m_filePath = settings.value("IniFilePath").toString();    //m_filePath = settings.value("LastTargetFilePath").toString();
      //m_imgPath = m_filePath+"\\"+m_fileName+".jpg";// add to insert image into excel
      QString strExcelFilePath = settings.value("ExcelFilePath").toString();  //*****************
      m_filePathExcel =  strExcelFilePath + "\\" + m_fileName +".xls";      //m_filePathExcel = m_filePath+"\\"+m_fileName+".xls";
      m_filePath += "\\"+m_fileName+".ini"; // mark
//      qDebug() << m_filePath;
//      qDebug() << m_filePathExcel;

      QString strExcelRepairTemp = settings.value("TemplateVersionRepair").toString();  //Repair_matrix_MK5  //************
//    Template for the Repair Sheet:
      m_filePathExcelRepairTmp = QDir::toNativeSeparators(QApplication::applicationDirPath()) + "\\" + strExcelRepairTemp + ".xlsx";
      QString m_robotNumber;
      QStringList parts = m_fileName.split("_");
      if(parts.size() > 1) {
        m_robotNumber = parts[1];
      }

      m_filePathExcelRepair =strExcelFilePath + "\\" + strExcelRepairTemp + "_"+ m_robotNumber +"_w.xlsx"; // strExcelFilePath + "\\" + strExcelRepairTemp + "_" + m_robotNumber +"_w.xlsx";

      QString strExcelMOMTemp = settings.value("TemplateVersionMOM").toString();  //MOM_ARR-NT  //************
//    Template for the MOM Sheet:
      m_filePathExcelMOMTmp = QDir::toNativeSeparators(QApplication::applicationDirPath()) + "\\" + strExcelMOMTemp + ".xlsx";
      strExcelMOMTemp.chop(6);
      m_filePathExcelMOM =strExcelFilePath + "\\" + strExcelMOMTemp + m_robotNumber +".xlsx"; // strExcelFilePath + "\\" + strExcelRepairTemp + "_" + m_robotNumber +"_w.xlsx";



      // for image path //ImageFilePath
      m_imgDefaultPath = settings.value("ImageFilePath").toString();
//      qDebug()<<"Defimgpath = "<<m_imgDefaultPath<<endl;
      progressSave(0);
//      setRadioButtonsIDsInGB4ARM();
//      setRadioButtonsIDsInGB4DM();
//      setRadioButtonsIDsInGB4ZT();
      if(QFile::exists(m_filePath))
      {
        getDataFromIni4Arm();
        getDataFromIni4DM();
        getDataFromIni4ZT();
        getDataFromIni4Repair();
        split3PartFromSN();
          // for test
//          getDataFromIni4ArmAnalyseData();
      }
      else
      {// need to clear data
          initVal();
//          qDebug()<<"After Init"<<endl;
          createNewIniFile();
//          qDebug()<<"Create"<<endl;
          // ANALYSE_NT2468_KW2316 >>> [0]:ANALYSE->Protocol; [1]:NT2468; [2]:KW2316
          split3PartFromSN();
      }
    //}
}

void CAnalyseData::getDataFromIni4Arm( void)
{
    getDataFromIni4ArmSN();
    getDataFromIni4ArmNCNR();
    getDataFromIni4ArmGrayBoxData();
    getDataFromIni4HDMotorType();
    getDataFromIni4ArmAdviceCauser();
    getDataFromIni4ArmOKUpgrade();
    getDataFromIni4ArmVacFlowVal();
    getDataFromIni4ArmGeoData();
    getDataFromIni4ArmTestsResults();
    getDataFromIni4ArmRepPosPAData();
    getDataFromIni4ArmAnalyseData();
    getDataFromIni4ArmEndDefectTiltData();
    getDataFromIni4ArmULVal();
    getDataFromIni4ArmComments();// to get ARM comments
}

void CAnalyseData::getDataFromIni4ArmRobTypeNo( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strValue;
    strValue = settings.value("robotType").toString();
    ui->lbRobotType->setText(strValue);
}
void CAnalyseData::getDataFromIni4ArmSN( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ROBOTINFO);
    QString strValue;
    strValue = settings.value("10/val").toString();
    ui->leARMSN->setText(strValue);
}

void CAnalyseData::getDataFromIni4ArmNCNR( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strValue;
    strValue = settings.value("robotType").toString();
    ui->lbRobotType->setText(strValue);

    strValue = "";
    strValue = settings.value("robotNo").toString();
    ui->leRobotTypeSN->setText(strValue);

    strValue = "";
    strValue = settings.value("robotSN").toString();
    ui->leRobotSN->setText(strValue);

    strValue = "";
    strValue = settings.value("old12NC").toString();
    ui->leOld12NC->setText(strValue);

    strValue = "";
    strValue = settings.value("new12NC").toString();
    ui->leNew12NC->setText(strValue);

    strValue = "";
    strValue = settings.value("repairNR").toString();
    ui->leRepairNRARM->setText(strValue);

    m_imgPath = "";
    strValue = "";
    strValue = settings.value("ImgFilePath").toString();
    ui->lbImagePath->setText(strValue);
    m_imgPath = strValue + "/";
//    qDebug()<<"imgPt = "<<m_imgPath<<endl;

    strValue = "";
    strValue = settings.value("BIDCode").toString();
    ui->leBIDCode->setText(strValue);

    strValue = "";
    strValue = settings.value("ImgFileName").toString();
    ui->lbImageName->setText(strValue);
    m_imgPath = m_imgPath + strValue;
//    qDebug()<<"ImgPath:"<<m_imgPath<<endl;


}

void CAnalyseData::getDataFromIni4ArmGrayBoxData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strValue;
    strValue = settings.value("grayboxSN").toString();
    ui->leGrayboxSN->setText(strValue);

    strValue = "";
    strValue = settings.value("firstDeliveryDate").toString();        
    ui->leFirstDeliveryDate->setText(strValue);

    strValue = "";
    strValue = settings.value("repairNo").toString();
    ui->leRepairNo->setText(strValue);

    strValue = "";
    strValue = settings.value("lastRepairDate").toString();
    ui->leLastRepairDate->setText(strValue);

    strValue = "";
    strValue = settings.value("firstDeliveryDateARM").toString();
    ui->leArmFirstDelivery->setText(strValue);

    strValue = "";
    strValue = settings.value("lastRepairDateARM").toString();
    ui->leArmLastRepair->setText(strValue);

    if(ui->leRobotSN->text() == "" || ui->leARMSN->text() == "" || ui->leRepairNRARM->text() == "" || ui->leGrayboxSN->text() == "" ||
       ui->leRepairNo->text() == "" || ui->leLastRepairDate->text() == "" || ui->leFirstDeliveryDate->text() == "")
        QMessageBox::critical(NULL, "Error", "Arm Data Incomplete!", QMessageBox::Yes, QMessageBox::Yes);
}

void CAnalyseData::getDataFromIni4HDMotorType( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "hdMotorType";
    int iType = settings.value(strInputName).toInt();

    switch (iType)
    {
      case V0:
        ui->rbHDMotorType_V0->setChecked( true);
      break;
      case V1:
        ui->rbHDMotorType_V1->setChecked( true);
      break;
      case DFV1:
        ui->rbHDMotorType_DFV1->setChecked( true);
      break;
    }
}

void CAnalyseData::getDataFromIni4ArmAdviceCauser( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "analyseAdviceARM";
    int iType = settings.value(strInputName).toInt();

    switch (iType)
    {
      case NFF:
        ui->rbNFFARM->setChecked( true);
      break;
      case Warrenty:
        ui->rbWarrentyARM->setChecked( true);
      break;
      case GoodWill:
        ui->rbGoodWillARM->setChecked( true);
      break;
      case WithCosts:
        ui->rbWithCostsARM->setChecked( true);
      break;
      case ScrapItem:
        ui->rbScrapItemARM->setChecked( true);
      break;
    }

    strInputName = "";
    strInputName = "analyseCauserARM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Customer:
        ui->rbCustomerARM->setChecked( true);
      break;
      case ASYS:
        ui->rbASYSARM->setChecked( true);
      break;
    }// switch()
}

void CAnalyseData::getDataFromIni4ArmOKUpgrade( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "isUpgradeARM";
    int iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->cbUpgradeARM->setChecked( false);
      break;
      case Test_NG:
        ui->cbUpgradeARM->setChecked( true);
      break;
    }// switch()

    strInputName = "";
    strInputName = "isARMOk";
    iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->rbARMIsOK->setChecked( true);
      break;
      case Test_NG:
        ui->rbRepairARM->setChecked( true);
      break;
    }// switch()
}

void CAnalyseData::getDataFromIni4ArmVacFlowVal( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO);

    QString strInputName = "18/val";
    QString strValue;

    strValue = settings.value(strInputName).toString();
    ui->leVacuumARM->setText(strValue);
    strInputName = "";
    strValue = "";
    strInputName = "19/val";
    strValue = settings.value(strInputName).toString();
    ui->leFlowARM->setText(strValue);
}

//float  CAnalyseData::mmMradCovert(int len, bool bMm2Mrad, float fVal)
//{
//    float fRst;
//    // bMm2Rad = true ==> mm2MRad, false==> mRad2Mm
//    // len = 525 or 100
//    //double dMrad= double(qAtan(fVal/len)*1000);
//    return fRst;
//}

QString CAnalyseData::calGeoMaxVal(QString strName)
{
  QSettings settings(m_filePath, QSettings::IniFormat);
  settings.beginGroup(BEGIN_ARMINFO);
  float fMax = 0.0;
  float fVal[4];
  float fVal1[4];
  float fValDiff[4];
  float fMrad;
  QString strRst = "";

  if(strName=="geoRx")
  {
      fVal[0] = settings.value("4/val").toFloat();
      fVal[1] = settings.value("5/val").toFloat();
      fVal[2] = settings.value("6/val").toFloat();
      //qDebug()<<"fVal[0]:"<<fVal[0]<<"fVal[1]:"<<fVal[1]<<"fVal[2]:"<<fVal[2];
      if(fVal[0] == 9999 || fVal[1] == 9999 || fVal[2] == 9999)
         strRst = "N/A";
      else
      {
          for(int i = 0; i < 3; i++)
          {
              if(fabs(fVal[i])>=fabs(fMax))
              {
                  fMax = fVal[i];
              }
          }
          //qDebug()<<"fMax:"<<fMax;
          strRst = QString::number(fMax, 'f', 3);
      }
      //qDebug() << "geoRx: " <<  strRst << endl;
  }else if(strName=="geoRy")
  {
      fVal[0] = settings.value("7/val").toFloat();
      fVal[1] = settings.value("8/val").toFloat();
      fVal[2] = settings.value("9/val").toFloat();
      //qDebug()<<"fVal[0]:"<<fVal[0]<<"fVal[1]:"<<fVal[1]<<"fVal[2]:"<<fVal[2];
      if(fVal[0] == 9999 || fVal[1] == 9999 || fVal[2] == 9999)
          strRst = "N/A";
      else
      {
          for(int j = 0; j < 3; j++)
          {
              if(fabs(fVal[j])>=fabs(fMax))
              {
                  fMax = fVal[j];
              }
          }
        //qDebug()<<"fMax:"<<fMax;
        strRst = QString::number(fMax, 'f', 3);
      }
      //qDebug() << "geoRy: " << strRst << endl;
  }else if(strName=="geoDeletaH")
  {
      fVal[0] = settings.value("29/val").toFloat();
      fVal[1] = settings.value("31/val").toFloat();
      //qDebug() << "fVal[0]: " << fVal[0] << "  fVal[1]: " << fVal[1];
      if(fVal[0] == 9999 || fVal[1] == 9999)
          strRst = "N/A";
      else
      {
        fMax = fVal[1] - fVal[0];
        strRst = QString::number(fMax, 'f', 3);
        //qDebug() << "fmax: "<< fMax;
      }
      //qDebug() << "geoDeltaH: " << strRst << endl;
  }else//geoRz
  {
      fVal[0] = settings.value("21/val").toFloat();
      fVal[1] = settings.value("23/val").toFloat();
      fVal[2] = settings.value("25/val").toFloat();
      fVal[3] = settings.value("32/val").toFloat();
      fVal1[0] = settings.value("22/val").toFloat();
      fVal1[1] = settings.value("24/val").toFloat();
      fVal1[2] = settings.value("26/val").toFloat();
      fVal1[3] = settings.value("33/val").toFloat();
      // need to convert to mrad then get the max.
      //qDebug()<<"fVal[0]:"<<fVal[0]<<"fVal[1]:"<<fVal[1]<<"fVal[2]:"<<fVal[2];
      //qDebug()<<"fVal1[0]:"<<fVal1[0]<<"fVal1[1]:"<<fVal1[1]<<"fVal1[2]:"<<fVal1[2];

      if(fVal[0] == 9999 || fVal[1] == 9999 || fVal[2] == 9999 || fVal[3] == 9999 || fVal1[0] == 9999 || fVal1[1] == 9999 || fVal1[2] == 9999 || fVal1[3] == 9999)
          strRst = "N/A";
      else
      {
          for(int k = 0; k < 4; k++)
          {
              fValDiff[k] = fVal1[k] - fVal[k];
              fMrad = float(qAtan(fValDiff[k]/TOOL_LENGTH1)*1000);
              if(fabs(fMrad)>=fabs(fMax))
              {
                  fMax = fMrad;
              }
          }
        //qDebug()<<"fMax:"<<fMax;
        strRst = QString::number(fMax, 'f', 3);
      }
      //qDebug() << "geoRZ: " << strRst << endl;
  }

//  strRst = QString::number(fMax, 'f', 3);
  return strRst;
}

void CAnalyseData::getDataFromIni4ArmGeoData( void)
{
//    QSettings settings(m_filePath,QSettings::IniFormat);
//    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "geoRz";
    QString strValue;

    strValue = calGeoMaxVal(strInputName);
    ui->leGeoRz->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "geoRx";
    strValue = calGeoMaxVal(strInputName);
    ui->leGeoRx->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "geoRy";
    strValue = calGeoMaxVal(strInputName);
    ui->leGeoRy->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "geoDeletaH";
    strValue = calGeoMaxVal(strInputName);
    ui->leGeoDelHeight->setText(strValue);
}

void CAnalyseData::getDataFromIni4ArmTestsResults( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);

    QString strInputName = "isEletricityOkARM";
    int iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->rbEletricityOKARM->setChecked( true);
        break;
      case Test_NG:
        ui->rbEletricityNOKARM->setChecked( true);
        break;
      case Test_NA:
        ui->rbEletricityNAARM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isDataTransOkARM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbDataTransOKARM->setChecked( true);
      break;
      case Test_NG:
        ui->rbDataTransNOKARM->setChecked( true);
      break;
      case Test_NA:
        ui->rbDataTransNAARM->setChecked( true);
      break;
    }// switch()

    strInputName = "";
    strInputName = "isSurOkARM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbSurfNotDamegeOKARM->setChecked( true);
        break;
      case Test_NG:
        ui->rbSurfNotDamegeNOKARM->setChecked( true);
        break;
      case Test_NA:
        ui->rbSurfNotDamegeNAARM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isFunOkARM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbFunTestOKARM->setChecked( true);
      break;
      case Test_NG:
        ui->rbFunTestNOKARM->setChecked( true);
      break;
      case Test_NA:
        ui->rbFunTestNAARM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isMagneticOkARM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbMagAttOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbMagAttNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbMagAttNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isGeoOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbGeoOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbGeoNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbGeoNA->setChecked( true);
        break;
    }// switch()
}

void CAnalyseData::getDataFromIni4ArmRepPosPAData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strInputName = "25/val";//"repPosPA_TH";
    QString strValue;

    strValue = settings.value(strInputName).toString();
    if(strValue != "N/A")
    {
        float fVal;
        fVal = strValue.toFloat()*1000;
        //qDebug()<<"PA-TH:"<<fVal<<endl;
        strValue = QString::number(fVal, 'f', 3);//fVal;//
    //    qDebug()<<"PA-TH:"<<fVal<<"strValue:"<<strValue<<endl;
        //strValue = settings.value(strInputName).toString();
    }
    ui->leRepPosPAR->setText(strValue);


    strInputName = "";
    strValue = "";
    strInputName = "26/val";    //"repPosPA_R";
    strValue = settings.value(strInputName).toString();
    if(strValue != "N/A")
    {
        float fVal;
        fVal = strValue.toFloat()*1000;
        //qDebug()<<"PA-R:"<<fVal<<endl;
        strValue = QString::number(fVal, 'f', 3);//fVal;//
    //    qDebug()<<"PA-R:"<<fVal<<"strValue:"<<strValue<<endl;
        //strValue = settings.value(strInputName).toString();
    }
    ui->leRepPosPATH->setText(strValue);
}

void CAnalyseData::getDataFromIni4ArmAnalyseData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_TESTERINFO);
    QString strInputName, strValue;
    QStringList strDateTmp;

    strInputName = "1/Tester";
    strValue = settings.value(strInputName).toString();
    ui->leAnalysePerformerARM->setText(strValue);
    ui->leAnalysePerformerDM->setText(strValue);
    ui->leAnalysePerformerZT->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "1/Time";
    strValue = settings.value(strInputName).toString();
    strDateTmp = strValue.split(" ");
//    for( int i = 0; i < strDateTmp.count(); i++)
//    {
//      qDebug()<<"strDateTmp:"<<strDateTmp[i]<<endl;
//    }
    ui->leAnalyseDateARM->setText(strDateTmp[0]);//strValue
    ui->leAnalyseDateDM->setText(strDateTmp[0]);
    ui->leAnalyseDateZT->setText(strDateTmp[0]);
}

void CAnalyseData::getDataFromIni4ArmEndDefectTiltData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO);
    QString strInputName = "34/val";//"rMinus208H4Val";
    int i = 0;
    float fGetVal= 0.0;
    float fValConvert= 0.0;
    // for Rx
    //m_strTiltValRx[0] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRx[i] = settings.value(strInputName).toString();
    m_strTiltValRx[i] = QString::number(fValConvert, 'f', 3);
    ui->leRMinus208Rx->setText(m_strTiltValRx[i]);

    i++;
    strInputName = "";
    strInputName = "35/val";
    //m_strTiltValRx[1] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRx[i] = settings.value(strInputName).toString();
    m_strTiltValRx[i] = QString::number(fValConvert, 'f', 3);
    ui->leR0Rx->setText(m_strTiltValRx[i]);

    i++;
    strInputName = "";
    strInputName = "36/val";
    //m_strTiltValRx[2] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRx[i] = settings.value(strInputName).toString();
    m_strTiltValRx[i] = QString::number(fValConvert, 'f', 3);
    ui->leR75Rx->setText(m_strTiltValRx[i]);

    i++;
    strInputName = "";
    strInputName = "4/val";
    //m_strTiltValRx[3] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRx[i] = settings.value(strInputName).toString();
    m_strTiltValRx[i] = QString::number(fValConvert, 'f', 3);
    ui->leR174Rx->setText(m_strTiltValRx[i]);

    i++;
    strInputName = "";
    strInputName = "5/val";
    //m_strTiltValRx[4] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRx[i] = settings.value(strInputName).toString();
    m_strTiltValRx[i] = QString::number(fValConvert, 'f', 3);
    ui->leR203Rx->setText(m_strTiltValRx[i]);

    i++;
    strInputName = "";
    strInputName = "6/val";
    //m_strTiltValRx[5] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRx[i] = settings.value(strInputName).toString();
    m_strTiltValRx[i] = QString::number(fValConvert, 'f', 3);
    ui->leR385Rx->setText(m_strTiltValRx[i]);
    // for Ry
    i = 0;
    strInputName = "";
    strInputName = "37/val";
    //m_strTiltValRy[0] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRy[i] = settings.value(strInputName).toString();
    m_strTiltValRy[i] = QString::number(fValConvert, 'f', 3);
    ui->leRMinus208Ry->setText(m_strTiltValRy[i]);

    i++;
    strInputName = "";
    strInputName = "38/val";
    //m_strTiltValRy[1] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRy[i] = settings.value(strInputName).toString();
    m_strTiltValRy[i] = QString::number(fValConvert, 'f', 3);
    ui->leR0Ry->setText(m_strTiltValRy[i]);

    i++;
    strInputName = "";
    strInputName = "39/val";
    //m_strTiltValRy[2] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRy[i] = settings.value(strInputName).toString();
    m_strTiltValRy[i] = QString::number(fValConvert, 'f', 3);
    ui->leR75Ry->setText(m_strTiltValRy[i]);

    i++;
    strInputName = "";
    strInputName = "7/val";
    //m_strTiltValRy[3] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRy[i] = settings.value(strInputName).toString();
    m_strTiltValRy[i] = QString::number(fValConvert, 'f', 3);
    ui->leR174Ry->setText(m_strTiltValRy[i]);

    i++;
    strInputName = "";
    strInputName = "8/val";
    //m_strTiltValRy[4] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRy[i] = settings.value(strInputName).toString();
    m_strTiltValRy[i] = QString::number(fValConvert, 'f', 3);
    ui->leR203Ry->setText(m_strTiltValRy[i]);

    i++;
    strInputName = "";
    strInputName = "9/val";
    //m_strTiltValRy[5] = settings.value(strInputName).toString();
    fGetVal = settings.value(strInputName).toFloat();
    fValConvert = qTan(fGetVal/1000)*TOOL_LENGTH1;
    //m_strTiltValRy[i] = settings.value(strInputName).toString();
    m_strTiltValRy[i] = QString::number(fValConvert, 'f', 3);
    ui->leR385Ry->setText(m_strTiltValRy[i]);

    // for H4
    i = 0;
    strInputName = "";
    strInputName = "40/val";
    //m_strTiltValH4[0] = settings.value(strInputName).toString();
    m_strTiltValH4[i] = settings.value(strInputName).toString();
    ui->leRMinus208H4->setText(m_strTiltValH4[i]);

    i++;
    strInputName = "";
    strInputName = "41/val";
    //m_strTiltValH4[1] = settings.value(strInputName).toString();
    m_strTiltValH4[i] = settings.value(strInputName).toString();
    ui->leR0H4->setText(m_strTiltValH4[i]);

    i++;
    strInputName = "";
    strInputName = "42/val";
    //m_strTiltValH4[2] = settings.value(strInputName).toString();
    m_strTiltValH4[i] = settings.value(strInputName).toString();
    ui->leR75H4->setText(m_strTiltValH4[i]);

    i++;
    strInputName = "";
    strInputName = "30/val";
    //m_strTiltValH4[3] = settings.value(strInputName).toString();
    m_strTiltValH4[i] = settings.value(strInputName).toString();
    ui->leR174H4->setText(m_strTiltValH4[i]);

    i++;
    strInputName = "";
    strInputName = "29/val";
    //m_strTiltValH4[4] = settings.value(strInputName).toString();
    m_strTiltValH4[i] = settings.value(strInputName).toString();
    ui->leR203H4->setText(m_strTiltValH4[i]);

    i++;
    strInputName = "";
    strInputName = "31/val";
    //m_strTiltValH4[5] = settings.value(strInputName).toString();
    m_strTiltValH4[i] = settings.value(strInputName).toString();
    ui->leR385H4->setText(m_strTiltValH4[i]);
}

void CAnalyseData::getDataFromIni4ArmULVal( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO);
    QString strInputName = "43/val";//"analyseUVal";
    QString strValue;

    strValue = settings.value(strInputName).toString();
    ui->leOAG->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "44/val";
    strValue = settings.value(strInputName).toString();
    ui->leUAG->setText(strValue);
}

void CAnalyseData::getDataFromIni4ArmComments( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strComments = settings.value("analyseCommentsARM").toString();
    ui->teARMComments->setText(strComments);

    //20230810 add
    m_RemarkImgPath_ARM = "";
    m_RemarkImgFileNameARM = "";

    QString strIniImgFilePath = settings.value("ARMRemarkImgFilePath").toString();
    QString strIniImgFileName = settings.value("ARMRemarkImgFileName").toString();

    if(strIniImgFilePath != "" && strIniImgFileName != "")    //Image data exists
    {
        m_RemarkImgPath_ARM = strIniImgFilePath;
        m_RemarkImgFileNameARM = strIniImgFileName;
        ui->lbl_ImgRemarkFileName_ARM->setText(strIniImgFileName);
    }else   //Image data don't exists
    {
        ui->lbl_ImgRemarkFileName_ARM->clear();
    }
}

void CAnalyseData::getDataFromIni4DM( void)
{

    getDataFromIni4DMSN();
    getDataFromIni4DMAdviceCauser();
    getDataFromIni4DMOKUpgrade();
    getDataFromIni4DMVacFlowVal();
    getDataFromIni4DMAngleData();
    getDataFromIni4DMTestsResults();
    getDataFromIni4DMCommutationData();
    getDataFromIni4DMZeroingPosData();
    //getDataFromIni4DMAnalyseData();
    getDataFromIni4DMComments();
    getDataFromIni4DMDeliveryRepairDate();
}

void CAnalyseData::getDataFromIni4DMSN( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    //settings.beginGroup(BEGIN_DMINFO);
    settings.beginGroup(BEGIN_ROBOTINFO);//"4/val";//
    QString strValue = "", strTmp;
    QStringList strListVal;
    strTmp = settings.value("4/val").toString();// get "ARD-140-BD-BA-016-20-47-8404-AP"
    strListVal = strTmp.split("-");
    if(strListVal.count()!=9)
    {
        QMessageBox::critical(NULL, "Error", "Drive Module Data Incomplete!", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }
    else
    {
        strValue = strListVal[4]+"-"+strListVal[5]+"-"+strListVal[6]+"-"+strListVal[7]+"-"+strListVal[8];
        ui->leDMSN->setText(strValue);
    }
}

void CAnalyseData::getDataFromIni4DMAdviceCauser( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strInputName = "analyseAdviceDM";
    int iType = settings.value(strInputName).toInt();

    switch (iType)
    {
      case NFF:
        ui->rbNFFDM->setChecked( true);
      break;
      case Warrenty:
        ui->rbWarrentyDM->setChecked( true);
      break;
      case GoodWill:
        ui->rbGoodWillDM->setChecked( true);
      break;
      case WithCosts:
        ui->rbWithCostsDM->setChecked( true);
      break;
      case ScrapItem:
        ui->rbScrapItemDM->setChecked( true);
      break;
    }

    strInputName = "";
    strInputName = "analyseCauserDM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Customer:
        ui->rbCustomerDM->setChecked( true);
      break;
      case ASYS:
        ui->rbASYSDM->setChecked( true);
      break;
    }// switch()
}

void CAnalyseData::getDataFromIni4DMOKUpgrade( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strInputName = "isUpgradeDM";
    int iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->cbUpgradeDM->setChecked( false);
      break;
      case Test_NG:
        ui->cbUpgradeDM->setChecked( true);
      break;
    }// switch()

    strInputName = "";
    strInputName = "isDMOk";
    iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->rbDMIsOK->setChecked( true);
      break;
      case Test_NG:
        ui->rbRepairDM->setChecked( true);
      break;
    }// switch()
}

void CAnalyseData::getDataFromIni4DMVacFlowVal( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);

    QString strInputName = "46/val";//"vacDM";
    QString strValue;

    strValue = settings.value(strInputName).toString();
    ui->leVacuumDM->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "45/val";//"flowDM";
    strValue = settings.value(strInputName).toString();
    ui->leFlowDM->setText(strValue);
}

void CAnalyseData::getDataFromIni4DMAngleData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strInputName;// = "ang90Val";
    QString strValue;
    float fMrad, fVal;

//    strValue = settings.value(strInputName).toString();
//    ui->le90DegVal->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "83/val";    //"ang180Val";
    strValue = settings.value(strInputName).toString();
    ui->le180DegVal_2->setText(strValue);

    if(strValue != "N/A" && isDigitStr(strValue) == true)
    {
        fVal = strValue.toFloat();
        fMrad = float(qAtan(fVal/TOOL_LENGTH)*1000);
        strValue = QString::number(fMrad, 'f', 3);
        ui->le180DegVal->setText(strValue);
    }else
        ui->le180DegVal->setText("N/A");

    strInputName = "";
    strValue = "";
    strInputName = "84/val";//"ang270Val";
    strValue = settings.value(strInputName).toString();
    ui->le270DegVal_2->setText(strValue);

    if(strValue != "N/A" && isDigitStr(strValue) == true)
    {
        fVal = strValue.toFloat();
        fMrad = float(qAtan(fVal/TOOL_LENGTH)*1000);
        strValue = QString::number(fMrad, 'f', 3);
        ui->le270DegVal->setText(strValue);
    }else
        ui->le270DegVal->setText("N/A");
}

void CAnalyseData::getDataFromIni4DMTestsResults( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);

    QString strInputName = "isMotorTHOk";
    int iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->rbTHMotorOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbTHMotorNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbTHMotorNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isMotorROk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbRMotorOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbRMotorNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbRMotorNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isGearTHOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbTHGearOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbTHGearNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbTHGearNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isGearROk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbRGearOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbRGearNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbRGearNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isTiltOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbTiltOK->setChecked( true);
      break;
      case Test_NG:
        ui->rbTiltNOK->setChecked( true);
      break;
      case Test_NA:
        ui->rbTiltNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEncTHOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbEncJumpTestTHOK->setChecked( true);
      break;
      case Test_NG:
        ui->rbEncJumpTestTHNOK->setChecked( true);
      break;
      case Test_NA:
        ui->rbEncJumpTestTHNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEncROk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbEncJumpTestROK->setChecked( true);
        break;
      case Test_NG:
        ui->rbEncJumpTestRNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbEncJumpTestRNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEletricityOkDM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbEletricityOKDM->setChecked( true);
        break;
      case Test_NG:
        ui->rbEletricityNOKDM->setChecked( true);
        break;
      case Test_NA:
        ui->rbEletricityNADM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isDataTransOkDM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbDataTransTestOKDM->setChecked( true);
        break;
      case Test_NG:
        ui->rbDataTransTestNOKDM->setChecked( true);
        break;
      case Test_NA:
        ui->rbDataTransTestNADM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isSurOkDM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbSurfNotDamegeOKDM->setChecked( true);
        break;
      case Test_NG:
        ui->rbSurfNotDamegeNOKDM->setChecked( true);
        break;
      case Test_NA:
        ui->rbSurfNotDamegeNADM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isFunOkDM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbFunTestOKDM->setChecked( true);
        break;
      case Test_NG:
        ui->rbFunTestNOKDM->setChecked( true);
        break;
      case Test_NA:
        ui->rbFunTestNADM->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isConductivityOkDM";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbConductivityChkOKDM->setChecked( true);
        break;
      case Test_NG:
        ui->rbConductivityChkNOKDM->setChecked( true);
        break;
      case Test_NA:
        ui->rbConductivityChkNADM->setChecked( true);
        break;
    }// switch()
}

void CAnalyseData::getDataFromIni4DMCommutationData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strInputName = "98/val";//"commutationTH";
    QString strValue;

    strValue = settings.value(strInputName).toString();
    if(strValue==NG_VALUE || strValue.contains("N/A"))
    {
        strValue = "N/A";
    }
    ui->leCommutationTH->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "101/val";//"commutationR";
    strValue = settings.value(strInputName).toString();
    if(strValue==NG_VALUE || strValue.contains("N/A"))
    {
        strValue = "N/A";
    }
    ui->leCommutationR->setText(strValue);
}

void CAnalyseData::getDataFromIni4DMZeroingPosData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strInputName = "zeroingPosTH";
    QString strValue;

    strValue = settings.value(strInputName).toString();
    ui->leZeroingPosTH->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "zeroingPosR";
    strValue = settings.value(strInputName).toString();
    ui->leZeroingPosR->setText(strValue);
}

void CAnalyseData::getDataFromIni4DMAnalyseData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_TESTERINFO);
    QString strInputName, strValue;

    strInputName = "1/Tester";
    strValue = settings.value(strInputName).toString();
    ui->leAnalysePerformerDM->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "1/Time";
    strValue = settings.value(strInputName).toString();
    ui->leAnalyseDateDM->setText(strValue);
}

void CAnalyseData::getDataFromIni4DMComments( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strComments = settings.value("analyseCommentsDM").toString();
    ui->teDMComments->setText(strComments);

    //20230810 add
    m_RemarkImgPath_DM = "";
    m_RemarkImgFileNameDM = "";

    QString strIniImgFilePath = settings.value("DMRemarkImgFilePath").toString();
    QString strIniImgFileName = settings.value("DMRemarkImgFileName").toString();

    if(strIniImgFilePath != "" && strIniImgFileName != "")    //Image data exists
    {
        m_RemarkImgPath_DM = strIniImgFilePath;
        m_RemarkImgFileNameDM = strIniImgFileName;
        ui->lbl_ImgRemarkFileName_DM->setText(strIniImgFileName);
    }else   //Image data don't exists
    {
        ui->lbl_ImgRemarkFileName_DM->clear();
    }
}

void CAnalyseData::getDataFromIni4DMDeliveryRepairDate()
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strValue;

    strValue = "";
    strValue = settings.value("firstDeliveryDateDM").toString();
    ui->leDMFirstDelivery->setText(strValue);

    strValue = "";
    strValue = settings.value("lastRepairDateDM").toString();
    ui->leDMLastRepair->setText(strValue);
}

void CAnalyseData::getDataFromIni4ZT( void)
{
    getDataFromIni4ZTSN();
    getDataFromIni4ZTAdviceCauser();
    getDataFromIni4ZTOKUpgrade();
    getDataFromIni4ZTMeasureData();
    getDataFromIni4ZTTestsResults();
    //getDataFromIni4ZTAnalyseData();
    getDataFromIni4ZTComments();
    getDataFromIni4ZTDeliveryRepairDate();
}

void CAnalyseData::getDataFromIni4ZTSN( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    //settings.beginGroup(BEGIN_ZTINFO);
    settings.beginGroup(BEGIN_ROBOTINFO);
    QString strValue, strTmp;
    QStringList strListVal;
    QString strZTLen;
    strTmp = settings.value("5/val").toString();// get "ARE-035-AA-AA-008-21-35-3768-AP"
    strListVal = strTmp.split("-");
    if(strListVal.count()!=9)
    {
        QMessageBox::critical(NULL, "Error", "Z-Drive Data Incomplete!", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }
    else
    {
        strZTLen = strListVal[1];
        ui->leZTLen->setText(strZTLen);
        strValue = strListVal[4]+"-"+strListVal[5]+"-"+strListVal[6]+"-"+strListVal[7]+"-"+strListVal[8];
        ui->leZTSN2->setText(strValue);
    }
}

void CAnalyseData::getDataFromIni4ZTAdviceCauser( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strInputName = "analyseAdviceZT";
    int iType = settings.value(strInputName).toInt();

    switch (iType)
    {
      case NFF:
        ui->rbNFFZT->setChecked( true);
      break;
      case Warrenty:
        ui->rbWarrentyZT->setChecked( true);
      break;
      case GoodWill:
        ui->rbGoodWillZT->setChecked( true);
      break;
      case WithCosts:
        ui->rbWithCostsZT->setChecked( true);
      break;
      case ScrapItem:
        ui->rbScrapItemZT->setChecked( true);
      break;
    }

    strInputName = "";
    strInputName = "analyseCauserZT";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Customer:
        ui->rbCustomerZT->setChecked( true);
      break;
      case ASYS:
        ui->rbASYSZT->setChecked( true);
      break;
    }// switch()
}

void CAnalyseData::getDataFromIni4ZTOKUpgrade( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strInputName = "isUpgradeZT";
    int iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->cbUpgradeZT->setChecked( false);
      break;
      case Test_NG:
        ui->cbUpgradeZT->setChecked( true);
      break;
    }// switch()

    strInputName = "";
    strInputName = "isZTOk";
    iType = settings.value(strInputName).toInt();

    switch(iType)
    {
      case Test_OK:
        ui->rbZTIsOK->setChecked( true);
      break;
      case Test_NG:
        ui->rbRepairZT->setChecked( true);
      break;
    }// switch()
}

void CAnalyseData::getDataFromIni4ZTMeasureData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    //settings.beginGroup(BEGIN_ZTINFO);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "robotType";
    QString strValue, strUpVal, strDNVal;
    strValue = settings.value(strInputName).toString();
    settings.endGroup();

    ui->leZUpFA->clear();
    ui->leZDownFA->clear();
    ui->leZUpSCARANT->clear();
    ui->leZDownSCARANT->clear();
    ui->leZUpNXT->clear();
    ui->leZDownNXT->clear();
    ui->leZUpSCARA->clear();
    ui->leZDownSCARA->clear();

    settings.beginGroup(BEGIN_DMINFO);
    strInputName = "";
    strInputName = "133/val";
    strUpVal = settings.value(strInputName).toString();

    strInputName = "";
    strInputName = "135/val";
    strDNVal = settings.value(strInputName).toString();
    if(strValue==ROBOTTYPE_DF)
    {
        ui->leZUpFA->setText(strUpVal);
        ui->leZDownFA->setText(strDNVal);
    }
    else if(strValue==ROBOTTYPE_NT)
    {
        ui->leZUpSCARANT->setText(strUpVal);
        ui->leZDownSCARANT->setText(strDNVal);
    }
    else if(strValue==ROBOTTYPE_NXT)
    {
        ui->leZUpNXT->setText(strUpVal);
        ui->leZDownNXT->setText(strDNVal);
    }
    else
    {
        ui->leZUpSCARA->setText(strUpVal);
        ui->leZDownSCARA->setText(strDNVal);
    }
}

void CAnalyseData::getDataFromIni4ZTTestsResults( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);

    QString strInputName = "isLVDTOk";
    int iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbLVDTOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbLVDTNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbLVDTNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isRefSnesorOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbRefSensorOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbRefSensorNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbRefSensorNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isMotorZOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbZMotorOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbZMotorNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbZMotorNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isRunningNoiseOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbRunNoiseOK->setChecked( true);
      break;
      case Test_NG:
        ui->rbRunNoiseNOK->setChecked( true);
      break;
      case Test_NA:
        ui->rbRunNoiseNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isCableOk";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbCableOK->setChecked( true);
        break;
      case Test_NG:
        ui->rbCableNOK->setChecked( true);
        break;
      case Test_NA:
        ui->rbCableNA->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isSurOkZT";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbSurfNotDamegeOKZT->setChecked( true);
        break;
      case Test_NG:
        ui->rbSurfNotDamegeNOKZT->setChecked( true);
        break;
      case Test_NA:
        ui->rbSurfNotDamegeNAZT->setChecked( true);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isConductivityOkZT";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Test_OK:
        ui->rbConductivityChkOKZT->setChecked( true);
        break;
      case Test_NG:
        ui->rbConductivityChkNOKZT->setChecked( true);
        break;
      case Test_NA:
        ui->rbConductivityChkNAZT->setChecked( true);
        break;
    }// switch()
}

void CAnalyseData::getDataFromIni4ZTAnalyseData( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_TESTERINFO);
    QString strInputName, strValue;

    strInputName = "1/Tester";
    strValue = settings.value(strInputName).toString();
    ui->leAnalysePerformerZT->setText(strValue);

    strInputName = "";
    strValue = "";
    strInputName = "1/Time";
    strValue = settings.value(strInputName).toString();
    ui->leAnalyseDateZT->setText(strValue);
}

void CAnalyseData::getDataFromIni4ZTComments( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strComments = settings.value("analyseCommentsZT").toString();
    ui->teZTComments->setText(strComments);

    //20230810 add
    m_RemarkImgPath_ZT = "";
    m_RemarkImgFileNameZT = "";

    QString strIniImgFilePath = settings.value("ZTRemarkImgFilePath").toString();
    QString strIniImgFileName = settings.value("ZTRemarkImgFileName").toString();

    if(strIniImgFilePath != "" && strIniImgFileName != "")    //Image data exists
    {
        m_RemarkImgPath_ZT = strIniImgFilePath;
        m_RemarkImgFileNameZT = strIniImgFileName;
        ui->lbl_ImgRemarkFileName_ZT->setText(strIniImgFileName);
    }else   //Image data don't exists
    {
        ui->lbl_ImgRemarkFileName_ZT->clear();
    }
}

void CAnalyseData::getDataFromIni4ZTDeliveryRepairDate()
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strValue;

    strValue = "";
    strValue = settings.value("firstDeliveryDateZT").toString();
    ui->leZTFirstDelivery->setText(strValue);

    strValue = "";
    strValue = settings.value("lastRepairDateZT").toString();
    ui->leZTLastRepair->setText(strValue);
}

void CAnalyseData::on_pbBack_clicked()
{
    displaySet( true);
}

//----------------- Save remarks of ARM to ini file (text and image) ---------------//
void CAnalyseData::getTextDocumentFromArm( void)
{
    // test to get the contents of text edit
    QTextBlock textBlock;
    QStringList strList;
    QString strToIniComments = "";
    m_ptdDocumentARM = ui->teARMComments->document();
    strList.clear();
    m_comments4Arm = "";
    for( textBlock = m_ptdDocumentARM->begin();textBlock!=m_ptdDocumentARM->end();textBlock = textBlock.next())
    {
        //qDebug()<<textBlock.text()<<endl;
        strList += textBlock.text();
    }
    strToIniComments = "";
    for( int i = 0; i < strList.count(); i++)
    {
      //qDebug()<<"i = "<<i<<", str = "<<strList[i]<<endl;
      if( strList[i] != "")
      {
        strToIniComments += strList[i]+"\n";
      }
    }
    m_comments4Arm = strToIniComments;

    // Write a value to the INI file
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    settings.setValue("analyseCommentsARM", strToIniComments);

    //20230810 add
    QString strInputName = "";
    strInputName = "ARMRemarkImgFilePath";

    if(ui->lbl_ImgRemarkFileName_ARM->text() == "")
    {
        m_RemarkImgPath_ARM = "";
        m_RemarkImgFileNameARM = "";
    }
    settings.setValue(strInputName, m_RemarkImgPath_ARM);

    strInputName = "";
    strInputName = "ARMRemarkImgFileName";
    settings.setValue(strInputName, m_RemarkImgFileNameARM);
}

void CAnalyseData::getTextDocumentFromDM( void)
{
    // test to get the contents of text edit
    QTextBlock textBlock;
    QStringList strList;
    QString strToIniComments = "";
    m_ptdDocumentDM = ui->teDMComments->document();
    strList.clear();
    m_comments4DM = "";
    for( textBlock = m_ptdDocumentDM->begin();textBlock!=m_ptdDocumentDM->end();textBlock = textBlock.next())
    {
        //qDebug()<<textBlock.text()<<endl;
        strList += textBlock.text();
    }
    strToIniComments = "";
    for( int i = 0; i < strList.count(); i++)
    {
      if( strList[i] != "")
      {
        strToIniComments += strList[i]+"\n";
      }
    }
    m_comments4DM = strToIniComments;

    // Write a value to the INI file
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    settings.setValue("analyseCommentsDM", strToIniComments);

    //20230810 add
    QString strInputName = "";
    strInputName = "DMRemarkImgFilePath";

    if(ui->lbl_ImgRemarkFileName_DM->text() == "")
    {
        m_RemarkImgPath_DM = "";
        m_RemarkImgFileNameDM = "";
    }
    settings.setValue(strInputName, m_RemarkImgPath_DM);

    strInputName = "";
    strInputName = "DMRemarkImgFileName";
    settings.setValue(strInputName, m_RemarkImgFileNameDM);
}

void CAnalyseData::getDMDeliveryRepairDate()
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strDeliveryDate1st_DM, strLastRepairDate_DM;

    strDeliveryDate1st_DM = ui->leDMFirstDelivery->text();
    settings.setValue("firstDeliveryDateDM", strDeliveryDate1st_DM);
    strLastRepairDate_DM = ui->leDMLastRepair->text();
    settings.setValue("lastRepairDateDM", strLastRepairDate_DM);
}

void CAnalyseData::getTextDocumentFromZT( void)
{
    // test to get the contents of text edit
    QTextBlock textBlock;
    QStringList strList;
    QString strToIniComments = "";
    m_ptdDocumentZT = ui->teZTComments->document();
    strList.clear();
    m_comments4ZT = "";
    for( textBlock = m_ptdDocumentZT->begin();textBlock!=m_ptdDocumentZT->end();textBlock = textBlock.next())
    {
        //qDebug()<<textBlock.text()<<endl;
        strList += textBlock.text();
    }
    strToIniComments = "";
    for( int i = 0; i < strList.count(); i++)
    {
      if( strList[i] != "")
      {
        strToIniComments += strList[i]+"\n";
      }
    }
    m_comments4ZT = strToIniComments;

    // Write a value to the INI file
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    settings.setValue("analyseCommentsZT", strToIniComments);

    //20230810 add
    QString strInputName = "";
    strInputName = "ZTRemarkImgFilePath";

    if(ui->lbl_ImgRemarkFileName_ZT->text() == "")
    {
        m_RemarkImgPath_ZT = "";
        m_RemarkImgFileNameZT = "";
    }
    settings.setValue(strInputName, m_RemarkImgPath_ZT);

    strInputName = "";
    strInputName = "ZTRemarkImgFileName";
    settings.setValue(strInputName, m_RemarkImgFileNameZT);
}

void CAnalyseData::getZTDeliveryRepairDate()
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strDeliveryDate1st_ZT, strLastRepairDate_ZT;

    strDeliveryDate1st_ZT = ui->leZTFirstDelivery->text();
    settings.setValue("firstDeliveryDateZT", strDeliveryDate1st_ZT);
    strLastRepairDate_ZT = ui->leZTLastRepair->text();
    settings.setValue("lastRepairDateZT", strLastRepairDate_ZT);
}

void CAnalyseData::getArmSNFromArm( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ROBOTINFO);
    QString strArmSN;
    QString strName = "10/val";

    strArmSN = ui->leARMSN->text();
    settings.setValue(strName, strArmSN);

}

void CAnalyseData::getImagePathFromArm( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strImgData = "";
    QString strInputName = "ImgFilePath";

    strImgData = ui->lbImagePath->text();
    settings.setValue(strInputName, strImgData);

    strInputName = "";
    strImgData = "";
    strInputName = "ImgFileName";
    strImgData = ui->lbImageName->text();
    settings.setValue(strInputName, strImgData);

    // 20230706 add
    strInputName = "";
    strImgData = "";
    strInputName = "BIDCode";
    strImgData = ui->leBIDCode->text();
    settings.setValue(strInputName, strImgData);
}

void CAnalyseData::getRobotTypeNoSNNCNRFromArm( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strRobotType, strRobotNo, strRobotSN, strOld12NC, strNew12NC, repairNR;
    QString strInputName = "robotType";

    strRobotType = ui->lbRobotType->text();
    settings.setValue(strInputName, strRobotType);
    strRobotNo = ui->leRobotTypeSN->text();
    settings.setValue("robotNo", strRobotNo);

    strRobotSN = ui->leRobotSN->text();
    settings.setValue("robotSN", strRobotSN);

    strOld12NC = ui->leOld12NC->text();
    settings.setValue("old12NC", strOld12NC);
    strNew12NC = ui->leNew12NC->text();
    settings.setValue("new12NC", strNew12NC);
    repairNR = ui->leRepairNRARM->text();
    settings.setValue("repairNR", repairNR);
}

void CAnalyseData::getGrayBoxDataFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strGrayboxSN, strDeliveryDate1st, strRepairNo, strLastRepairDate;
    QString strDeliveryDate1st_ARM, strLastRepairDate_ARM;

    strGrayboxSN = ui->leGrayboxSN->text();
    settings.setValue("grayboxSN", strGrayboxSN);
    strDeliveryDate1st = ui->leFirstDeliveryDate->text();
    settings.setValue("firstDeliveryDate", strDeliveryDate1st);
    strRepairNo = ui->leRepairNo->text();
    settings.setValue("repairNo", strRepairNo);
    strLastRepairDate = ui->leLastRepairDate->text();
    settings.setValue("lastRepairDate", strLastRepairDate);

    strDeliveryDate1st_ARM = ui->leArmFirstDelivery->text();
    settings.setValue("firstDeliveryDateARM", strDeliveryDate1st_ARM);
    strLastRepairDate_ARM = ui->leArmLastRepair->text();
    settings.setValue("lastRepairDateARM", strLastRepairDate_ARM);
}

void CAnalyseData::getHDMotorTypeFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "hdMotorTpye";
    switch( m_pgbHDMotorType->checkedId())
    {
      case V0:
        settings.setValue(strInputName, V0);
      break;
      case V1:
        settings.setValue(strInputName, V1);
      break;
      case DFV1:
        settings.setValue(strInputName, DFV1);
      break;
    }
}

void CAnalyseData::getAdviceCauserFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "analyseAdviceARM";
    switch( m_pgbAnalyseAdvARM->checkedId())
    {
      case NFF:
        settings.setValue(strInputName, NFF);
      break;
      case Warrenty:
        settings.setValue(strInputName, Warrenty);
      break;
      case GoodWill:
        settings.setValue(strInputName, GoodWill);
      break;
      case WithCosts:
        settings.setValue(strInputName, WithCosts);
      break;
      case ScrapItem:
        settings.setValue(strInputName, ScrapItem);
      break;
    }// switch()
    strInputName = "";
    strInputName = "analyseCauserARM";
    switch( m_pgbCauserARM->checkedId())
    {
      case Customer:
        settings.setValue(strInputName, Customer);
      break;
      case ASYS:
        settings.setValue(strInputName, Warrenty);
      break;
    }// switch()
}

void CAnalyseData::getArmOKUpgradeFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strInputName = "isUpgradeARM";
    if(ui->cbUpgradeARM->isChecked())
    {
        settings.setValue(strInputName, Test_NG);
    }
    else
    {
        settings.setValue(strInputName, Test_OK);
    }
    strInputName = "";
    strInputName = "isARMOk";
    switch( m_pgbRepairARMUpgradeChk->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
      break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
      break;
    }// switch()
}

void CAnalyseData::getVacFlowValFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO);

    QString strVacARM, strFlowARM;
    strVacARM = ui->leVacuumARM->text();
    settings.setValue("18/val", strVacARM);
    strFlowARM = ui->leFlowARM->text();
    settings.setValue("19/val", strFlowARM);
}

void CAnalyseData::getULValFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO);
    QString strUVal, strLVal;

    strUVal = ui->leOAG->text();
    settings.setValue("43/val", strUVal);
    strLVal = ui->leUAG->text();
    settings.setValue("44/val", strLVal);
}

void CAnalyseData::getGeoDataFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);
    QString strGeoRz, strGeoRx, strGeoRy, strGeoDeletaH;
    strGeoRz = ui->leGeoRz->text();
    settings.setValue("geoRz", strGeoRz);
    strGeoRx = ui->leGeoRx->text();
    settings.setValue("geoRx", strGeoRx);
    strGeoRy = ui->leGeoRy->text();
    settings.setValue("geoRy", strGeoRy);
    strGeoDeletaH = ui->leGeoDelHeight->text();
    settings.setValue("geoDeletaH", strGeoDeletaH);
}

void CAnalyseData::getTestsResultsFromARM( void)        //Modify
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO1);

    QString strInputName = "isGeoOk";
    switch( m_pgbGeoChkARM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:                             //new code*****************************************
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEletricityOkARM";
    switch( m_pgbEleChkARM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:                             //new code*************************************
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isDataTransOkARM";
    switch( m_pgbDataTransChkARM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:                               //new code************************************
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isSurOkARM";
    switch( m_pgbSurfaceDamegeChkARM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:                             //new code***************************************
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isFunOkARM";
    switch( m_pgbFunChkARM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:                                 //new code*****************************************
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isMagneticOkARM";
    switch( m_pgbMagAttChkARM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);   //new code*****************************
        break;
    }// switch()
}

void CAnalyseData::getRepPosPADataFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strRepPosPA_R, strRepPosPA_TH;
    float fPosR, fPosTH;

    if(ui->leRepPosPAR->text() == "N/A" || isDigitStr(ui->leRepPosPAR->text()) == false)
        strRepPosPA_R = "N/A";
    else
    {
        //strRepPosPA_R = ui->leRepPosPAR->text().toFloat();
        fPosR = ui->leRepPosPAR->text().toFloat()/1000;
        strRepPosPA_R = QString::number(fPosR, 'f', 3);
    }
    settings.setValue("25/val", strRepPosPA_R);

    if(ui->leRepPosPATH->text() == "N/A" || isDigitStr(ui->leRepPosPATH->text()) == false)
        strRepPosPA_TH = "N/A";
    else
    {
        //strRepPosPA_TH = ui->leRepPosPATH->text();
        fPosTH = ui->leRepPosPATH->text().toFloat()/1000;
        strRepPosPA_TH = QString::number(fPosTH, 'f', 3);
    }
    settings.setValue("26/val", strRepPosPA_TH);
}

void CAnalyseData::getAnalyseDataFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_TESTERINFO);
    QString strName;
    QString strAnalyserARM, strAnalyseDateARM;
    strAnalyserARM = ui->leAnalysePerformerARM->text();
    strName = "1/Tester";
    settings.setValue(strName, strAnalyserARM);
    strName = "";
    strName = "1/Time";
    strAnalyseDateARM = ui->leAnalyseDateARM->text();
    settings.setValue(strName, strAnalyseDateARM);
}

void CAnalyseData::getEndDefectTiltDataFromARM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMINFO);

    QString strMinus208H4Val, strMinus208RxVal, strMinus208RyVal;
    QString strZeroH4Val, strZeroRxVal, strZeroRyVal;
    QString str75H4Val, str75RxVal, str75RyVal;
    QString str174H4Val, str174RxVal, str174RyVal;
    QString str203H4Val, str203RxVal, str203RyVal;
    QString str385H4Val, str385RxVal, str385RyVal;
//    strMinus208H4Val = ui->leRMinus208H4->text();
//    settings.setValue("rMinus208H4Val", strMinus208H4Val);
//    strMinus208RxVal = ui->leRMinus208Rx->text();
//    settings.setValue("rMinus208RxVal", strMinus208RxVal);
//    strMinus208RyVal = ui->leRMinus208Ry->text();
//    settings.setValue("rMinus208RyVal", strMinus208RyVal);

//    strZeroH4Val = ui->leR0H4->text();
//    settings.setValue("rZeroH4Val", strZeroH4Val);
//    strZeroRxVal = ui->leR0Rx->text();
//    settings.setValue("rZeroRxVal", strZeroRxVal);
//    strZeroRyVal = ui->leR0Ry->text();
//    settings.setValue("rZeroRyVal", strZeroRyVal);

//    str75H4Val = ui->leR75H4->text();
//    settings.setValue("r75H4Val", str75H4Val);
//    str75RxVal = ui->leR75Rx->text();
//    settings.setValue("r75RxVal", str75RxVal);
//    str75RyVal = ui->leR75Ry->text();
//    settings.setValue("r75RyVal", str75RyVal);

//    str174H4Val = ui->leR174H4->text();
//    settings.setValue("r174H4Val", str174H4Val);
//    str174RxVal = ui->leR174Rx->text();
//    settings.setValue("r174RxVal", str174RxVal);
//    str174RyVal = ui->leR174Ry->text();
//    settings.setValue("r174RyVal", str174RyVal);

//    str203H4Val = ui->leR203H4->text();
//    settings.setValue("r203H4Val", str203H4Val);
//    str203RxVal = ui->leR203Rx->text();
//    settings.setValue("r203RxVal", str203RxVal);
//    str203RyVal = ui->leR203Ry->text();
//    settings.setValue("r203RyVal", str203RyVal);

//    str385H4Val = ui->leR385H4->text();
//    settings.setValue("r385H4Val", str385H4Val);
//    str385RxVal = ui->leR385Rx->text();
//    settings.setValue("r385RxVal", str385RxVal);
//    str385RyVal = ui->leR385Ry->text();
//    settings.setValue("r385RyVal", str385RyVal);
}

void CAnalyseData::getDMSNFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    //settings.beginGroup(BEGIN_DMINFO);
    settings.beginGroup(BEGIN_ROBOTINFO);
    QString strDMSNPre, strDMSN, strDMAll;
    strDMSNPre = "ARD-140-BD-BA-";
    strDMSN = ui->leDMSN->text();
    strDMAll = strDMSNPre + strDMSN;
    settings.setValue("4/val", strDMAll);
}

void CAnalyseData::getAdviceCauserFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);

    QString strInputName = "analyseAdviceDM";
    switch( m_pgbAnalyseAdvDM->checkedId())
    {
      case NFF:
        settings.setValue(strInputName, NFF);
      break;
      case Warrenty:
        settings.setValue(strInputName, Warrenty);
      break;
      case GoodWill:
        settings.setValue(strInputName, GoodWill);
      break;
      case WithCosts:
        settings.setValue(strInputName, WithCosts);
      break;
      case ScrapItem:
        settings.setValue(strInputName, ScrapItem);
      break;
    }// switch()
    strInputName = "";
    strInputName = "analyseCauserDM";
    switch( m_pgbCauserDM->checkedId())
    {
      case Customer:
        settings.setValue(strInputName, Customer);
      break;
      case ASYS:
        settings.setValue(strInputName, Warrenty);
      break;
    }// switch()
}

void CAnalyseData::getDMOKUpgradeFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strInputName = "isUpgradeDM";
    if(ui->cbUpgradeDM->isChecked())
    {
        settings.setValue(strInputName, Test_NG);
    }
    else
    {
        settings.setValue(strInputName, Test_OK);
    }
    strInputName = "";
    strInputName = "isDMOk";
    switch( m_pgbRepairDMUpgradeChk->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
      break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
      break;
    }// switch()
}

void CAnalyseData::getVacFlowValFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);

    QString strVacDM, strFlowDM;
    strVacDM = ui->leVacuumDM->text();
    settings.setValue("46/val", strVacDM);
    strFlowDM = ui->leFlowDM->text();
    settings.setValue("45/val", strFlowDM);
}

void CAnalyseData::getAngleDataUnitMradFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);
    QString strAng180, strAng270;

    strAng180 = ui->le180DegVal->text();
    settings.setValue("ang180Val", strAng180);
    strAng270 = ui->le270DegVal->text();
    settings.setValue("ang270Val", strAng270);
}

void CAnalyseData::getAngleDataFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strAng180, strAng270;

    strAng180 = ui->le180DegVal_2->text();
    settings.setValue("83/val", strAng180);
    strAng270 = ui->le270DegVal_2->text();
    settings.setValue("84/val", strAng270);
}

void CAnalyseData::getTestsResultsFromDM( void)             //Modify
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);

    QString strInputName = "isMotorTHOk";
    switch( m_pgbTHMotorChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isMotorROk";
    switch( m_pgbRMotorChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isGearTHOk";
    switch( m_pgbTHGearChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isGearROk";
    switch( m_pgbRGearChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isTiltOk";
    switch( m_pgbTiltChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEncTHOk";
    switch( m_pgbEncJumpTHChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEncROk";
    switch( m_pgbEncJumpRChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isEletricityOkDM";
    switch( m_pgbEleChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isDataTransOkDM";
    switch( m_pgbDataTransChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isSurOkDM";
    switch( m_pgbSurfaceDamegeChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isFunOkDM";
    switch( m_pgbFunChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isConductivityOkDM";
    switch( m_pgbConductivityChkDM->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()
}

void CAnalyseData::getCommutationDataFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);

    QString strComTH, strComR;
    strComTH = ui->leCommutationTH->text();
    if(strComTH==NG_VALUE || strComTH.contains("N/A"))
    {
        strComTH = "N/A";
    }
    settings.setValue("98/val", strComTH);
    strComR = ui->leCommutationR->text();
    if(strComR==NG_VALUE|| strComR.contains("N/A"))
    {
        strComR = "N/A";
    }
    settings.setValue("101/val", strComR);
}

void CAnalyseData::getZeroingPosDataFromDM( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO1);

    QString strZeroPosTH, strZeroPosR;
    strZeroPosTH = ui->leZeroingPosTH->text();
    settings.setValue("zeroingPosTH", strZeroPosTH);
    strZeroPosR = ui->leZeroingPosR->text();
    settings.setValue("zeroingPosR", strZeroPosR);
}

void CAnalyseData::getAnalyseDataFromDM( void)
{// don't need
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strAnalyserDM, strAnalyseDateDM;
    strAnalyserDM = ui->leAnalysePerformerDM->text();
    settings.setValue("analyserDM", strAnalyserDM);
    strAnalyseDateDM = ui->leAnalyseDateDM->text();
    settings.setValue("analyseDateDM", strAnalyseDateDM);

//    QString strInputName = "isAnalyseOkDM";
//    switch( m_pgbAnalyseChkDM->checkedId())
//    {
//      case Test_OK:
//        settings.setValue(strInputName, Test_OK);
//      break;
//      case Test_NG:
//        settings.setValue(strInputName, Test_NG);
//      break;
//    }// switch()
}

void CAnalyseData::getZTSNFromZT( void)
{//cbZTSN
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ROBOTINFO);
    QString strDMSNPre, strZTLen, strDMSNMid, strDMSN, strDMSNAll;
    strDMSNPre = "ARE-";
    strZTLen = ui->leZTLen->text();
    strDMSNMid = "-AA-AA-";
    strDMSN = ui->leZTSN2->text();
    strDMSNAll = strDMSNPre + strZTLen + strDMSNMid + strDMSN;

    settings.setValue("5/val", strDMSNAll);
}

void CAnalyseData::getAdviceCauserFromZT( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);

    QString strInputName = "analyseAdviceZT";
    switch( m_pgbAnalyseAdvZT->checkedId())
    {
      case NFF:
        settings.setValue(strInputName, NFF);
      break;
      case Warrenty:
        settings.setValue(strInputName, Warrenty);
      break;
      case GoodWill:
        settings.setValue(strInputName, GoodWill);
      break;
      case WithCosts:
        settings.setValue(strInputName, WithCosts);
      break;
      case ScrapItem:
        settings.setValue(strInputName, ScrapItem);
      break;
    }// switch()

    strInputName = "";
    strInputName = "analyseCauserZT";
    switch( m_pgbCauserZT->checkedId())
    {
      case Customer:
        settings.setValue(strInputName, Customer);
      break;
      case ASYS:
        settings.setValue(strInputName, Warrenty);
      break;
    }// switch()
}

void CAnalyseData::getZTOKUpgradeFromZT( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strInputName = "isUpgradeZT";
    if(ui->cbUpgradeZT->isChecked())
    {
        settings.setValue(strInputName, Test_NG);
    }
    else
    {
        settings.setValue(strInputName, Test_OK);
    }
    strInputName = "";
    strInputName = "isZTOk";
    switch( m_pgbRepairZTUpgradeChk->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
      break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
      break;
    }// switch()
}

void CAnalyseData::getMeasureDataFromZT( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMINFO);
    QString strUpVal, strDNVal, strRobotType;
    // need to check Type then to get value
    strRobotType = ui->lbRobotType->text();
    if(strRobotType==ROBOTTYPE_DF)
    {
        strUpVal = ui->leZUpFA->text();
        strDNVal = ui->leZDownFA->text();
    }
    else if(strRobotType==ROBOTTYPE_NT)
    {
        strUpVal = ui->leZUpSCARANT->text();
        strDNVal = ui->leZDownSCARANT->text();
    }
    else if(strRobotType==ROBOTTYPE_NXT)
    {
        strUpVal = ui->leZUpNXT->text();
        strDNVal = ui->leZDownNXT->text();
    }
    else
    {
        strUpVal = ui->leZUpSCARA->text();
        strDNVal = ui->leZDownSCARA->text();
    }
    settings.setValue("133/val", strUpVal);
    settings.setValue("135/val", strDNVal);
}

void CAnalyseData::getTestsResultsFromZT( void)     //modify
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);

    QString strInputName = "isLVDTOk";
    switch( m_pgbLVDTZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isRefSnesorOk";
    switch( m_pgbRefSenZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
       settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isMotorZOk";
    switch( m_pgbZMotorZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
       settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isRunningNoiseOk";
    switch( m_pgbRunningNoiseZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isCableOk";
    switch( m_pgbCableZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isSurOkZT";
    switch( m_pgbSurfaceDamegeChkZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
         settings.setValue(strInputName, Test_NA);
        break;
    }// switch()

    strInputName = "";
    strInputName = "isConductivityOkZT";
    switch( m_pgbConductivityChkZT->checkedId())
    {
      case Test_OK:
        settings.setValue(strInputName, Test_OK);
        break;
      case Test_NG:
        settings.setValue(strInputName, Test_NG);
        break;
      case Test_NA:
        settings.setValue(strInputName, Test_NA);
        break;
    }// switch()
}

void CAnalyseData::getAnalyseDataFromZT( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTINFO);
    QString strAnalyserZT, strAnalyseDateZT;
    strAnalyserZT = ui->leAnalysePerformerZT->text();
    settings.setValue("analyserZT", strAnalyserZT);
    strAnalyseDateZT = ui->leAnalyseDateZT->text();
    settings.setValue("analyseDateZT", strAnalyseDateZT);
}

void CAnalyseData::on_pbSaveIni_clicked()
{
    qDebug() << "mFilePath" << m_filePath;
    ui->lbStatus->setText("Creating Ini file");
    // for Arm
    getDataFromGUIArm();
    // for DM
    getDataFromGUIDM();
    // for ZT
    getDataFromGUIZT();

    // others
    getArmSNFromArm();//armSN
    getDataFromRepair();
    qDebug() << ".ini file created";
    ui->lbStatus->setText("Ini file was created");
}


void CAnalyseData::getDataFromGUIArm(void)
{

    getImagePathFromArm();
    getRobotTypeNoSNNCNRFromArm();// get Robot type and no./robotSN/old12NC/new12NC/repairNR
    getGrayBoxDataFromARM();// get grayboxSN/firstDeliveryDate/repairNo/lastRepairDate
    getHDMotorTypeFromARM();// get hdMotorType
    getAdviceCauserFromARM();// get analyseAdviceARM/analyseCauserARM
    getArmOKUpgradeFromARM();// get isARMOk/isUpgradeARM
    getVacFlowValFromARM();// get vacARM/flowARM
    getGeoDataFromARM();// get geoRz/geoRx/geoRy/geoDeletaH/isGeoOk
    getTestsResultsFromARM();// get isEletricityOkARM/isDataTransOkARM/isSurOkARM/isFunOkARM/isMagneticOkARM
    getRepPosPADataFromARM();// get repPosPA_R/repPosPA_TH
    getAnalyseDataFromARM();// get analyserARM/analyseDateARM/isAnalyseOkARM
    getTextDocumentFromArm();// get Arm comments
    //getEndDefectTiltDataFromARM();// get rMinus208H4Val/rMinus208RxVal/rMinus208RyVal/
                                  //     rZeroH4Val/rZeroRxVal/rZeroRyVal/
                                  //     r75H4Val/r75RxVal/r75RyVal/
                                  //     r174H4Val/r174RxVal/r174RyVal/
                                  //     r203H4Val/r203RxVal/r203RyVal/
                                  //     r385H4Val/r385RxVal/r385RyVal
    getULValFromARM();// get analyseUVal/analyseLVal
}

void CAnalyseData::getDataFromGUIDM(void)
{
    getDMSNFromDM();// get DMSN
    getAdviceCauserFromDM();// get analyseAdviceDM/analyseCauserDM
    getDMOKUpgradeFromDM();// get isDMOk/isUpgradeDM
    getVacFlowValFromDM();// get vacDM/flowDM
    getAngleDataFromDM();// get ang90Val/ang180Val/ang270Val
    getTestsResultsFromDM();// get isMotorTHOk/isMotorROk/isGearTHOk/isGearROk/itTiltOk/isEncTHOk/isEncROk/isEletricityOkDM/isDataTransOkDM/isSurOkDM/isFunOkDM/isConductivityOkDM
    getCommutationDataFromDM();// get commutationTH/commutationR
    getZeroingPosDataFromDM();// get zeroingPosTH/zeroingPosR
    //getAnalyseDataFromDM();// get analyserDM/analyseDateDM/isAnalyseOkDM >>> not need
    getTextDocumentFromDM();// get analyseCommentsDM
    getDMDeliveryRepairDate();
}

void CAnalyseData::getDataFromGUIZT(void)
{
    getZTSNFromZT();// get ZTSN/ZTSN2
    getAdviceCauserFromZT();// get analyseAdviceZT/analyseCauserZT
    getZTOKUpgradeFromZT();// get isZTOk/isUpgradeZT
    getMeasureDataFromZT();// get zUpSCARAVal/zDNSCARAVal/zUpSCARANTVal/zDNSCARANTVal/zUpFAVal/zDNFAVal/zUpNXTVal/zDNNXTVal
    getTestsResultsFromZT();// get isLVDTOk/isRefSnesorOk/isMotorZOk/isRunningNoiseOk/isCableOk/isSurOkZT/isConductivityOkZT
    //getAnalyseDataFromZT();// get analyserZT/analyseDateZT/isAnalyseOkZT // not need
    getTextDocumentFromZT();// get analyseCommentsZT
    getZTDeliveryRepairDate();
}

void CAnalyseData::setRadioButtonsIDsInGB4ARM( void)
{
    // set Robot type group
//    m_pgbRobotTypeARM = new QButtonGroup( this);
//    m_pgbRobotTypeARM->addButton( ui->rbAAR, 0);
//    m_pgbRobotTypeARM->addButton( ui->rbAARNT, 1);
//    m_pgbRobotTypeARM->addButton( ui->rbNXT, 2);
//    ui->rbAAR->setChecked( true);
//    if(m_pgbAnalyseAdvARM!=nullptr)
//    {
//        delete m_pgbAnalyseAdvARM;
//        delete m_pgbCauserARM;
//        delete m_pgbRepairARMUpgradeChk;
//        delete m_pgbSurfaceDamegeChkARM;
//        delete m_pgbMagAttChkARM;
//        delete m_pgbEleChkARM;
//        delete m_pgbGeoChkARM;
//        delete m_pgbFunChkARM;
//        delete m_pgbDataTransChkARM;
//    }
    m_pgbHDMotorType = new QButtonGroup( this );
    m_pgbHDMotorType->addButton(ui->rbHDMotorType_V0, V0);
    m_pgbHDMotorType->addButton(ui->rbHDMotorType_V1, V1);
    m_pgbHDMotorType->addButton(ui->rbHDMotorType_DFV1, DFV1);


    // set Analyse Advice group
    m_pgbAnalyseAdvARM = new QButtonGroup( this);
    m_pgbAnalyseAdvARM->addButton(ui->rbNFFARM, 0);
    m_pgbAnalyseAdvARM->addButton(ui->rbWarrentyARM, 1);
    m_pgbAnalyseAdvARM->addButton(ui->rbGoodWillARM, 2);
    m_pgbAnalyseAdvARM->addButton(ui->rbWithCostsARM, 3);
    m_pgbAnalyseAdvARM->addButton(ui->rbScrapItemARM, 4);
    //ui->rbNFFARM->setChecked(true);

    // set Causer group
    m_pgbCauserARM = new QButtonGroup( this);
    m_pgbCauserARM->addButton( ui->rbCustomerARM, 0);
    m_pgbCauserARM->addButton( ui->rbASYSARM, 1);
    //ui->rbCustomerARM->setChecked( true);

    // set ARM repair or not
    m_pgbRepairARMUpgradeChk = new QButtonGroup( this);
    m_pgbRepairARMUpgradeChk->addButton( ui->rbRepairARM, 1);
    m_pgbRepairARMUpgradeChk->addButton( ui->rbARMIsOK, 0);
    //ui->rbRepairARM->setChecked( true);

    // set ARM Analyse result
//    m_pgbAnalyseChk = new QButtonGroup( this);
//    m_pgbAnalyseChk->addButton( ui->rbARMAnalyseOK, 0);
//    m_pgbAnalyseChk->addButton( ui->rbARMAnalyseNOK, 1);
//    ui->rbARMAnalyseNOK->setChecked( true);

    // set ARM surface damage result
    m_pgbSurfaceDamegeChkARM = new QButtonGroup( this);
    m_pgbSurfaceDamegeChkARM->addButton( ui->rbSurfNotDamegeOKARM, 0);
    m_pgbSurfaceDamegeChkARM->addButton( ui->rbSurfNotDamegeNOKARM, 1);
    m_pgbSurfaceDamegeChkARM->addButton( ui->rbSurfNotDamegeNAARM, 2);  //new code*****************
    //ui->rbSurfNotDamegeNOKARM->setChecked( true);

    // set ARM Magnetic result
    m_pgbMagAttChkARM = new QButtonGroup( this);
    m_pgbMagAttChkARM->addButton( ui->rbMagAttOK, 0);
    m_pgbMagAttChkARM->addButton( ui->rbMagAttNOK, 1);
    m_pgbMagAttChkARM->addButton( ui->rbMagAttNA, 2);  //new code*****************
    //ui->rbMagAttNOK->setChecked( true);

    // set ARM Electricity result
    m_pgbEleChkARM = new QButtonGroup( this);
    m_pgbEleChkARM->addButton( ui->rbEletricityOKARM, 0);
    m_pgbEleChkARM->addButton( ui->rbEletricityNOKARM, 1);
    m_pgbEleChkARM->addButton( ui->rbEletricityNAARM, 2);  //new code*****************
    //ui->rbEletricityNOKARM->setChecked( true);

    // set ARM GEO result
    m_pgbGeoChkARM = new QButtonGroup( this);
    m_pgbGeoChkARM->addButton( ui->rbGeoOK, 0);
    m_pgbGeoChkARM->addButton( ui->rbGeoNOK, 1);
    m_pgbGeoChkARM->addButton( ui->rbGeoNA, 2);  //new code*****************
   //ui->rbGeoNOK->setChecked( true);

    // set ARM function result
    m_pgbFunChkARM = new QButtonGroup( this);
    m_pgbFunChkARM->addButton( ui->rbFunTestOKARM, 0);
    m_pgbFunChkARM->addButton( ui->rbFunTestNOKARM, 1);
    m_pgbFunChkARM->addButton( ui->rbFunTestNAARM, 2);  //new code*****************
    //ui->rbFunTestNOKARM->setChecked( true);

    // set ARM Data transfer result
    m_pgbDataTransChkARM = new QButtonGroup( this);
    m_pgbDataTransChkARM->addButton( ui->rbDataTransOKARM, 0);
    m_pgbDataTransChkARM->addButton( ui->rbDataTransNOKARM, 1);
    m_pgbDataTransChkARM->addButton( ui->rbDataTransNAARM, 2);  //new code*****************
    //ui->rbDataTransNOKARM->setChecked( true);
}

// for DM
void CAnalyseData::setRadioButtonsIDsInGB4DM( void)
{
//    if(m_pgbAnalyseAdvDM!=nullptr)
//    {
//        delete m_pgbAnalyseAdvDM;
//        delete m_pgbCauserDM;
//        delete m_pgbRepairDMUpgradeChk;
//        delete m_pgbSurfaceDamegeChkDM;
//        delete m_pgbEleChkDM;
//        delete m_pgbFunChkDM;
//        delete m_pgbDataTransChkDM;
//        delete m_pgbTHMotorChkDM;
//        delete m_pgbRMotorChkDM;
//        delete m_pgbTHGearChkDM;
//        delete m_pgbRGearChkDM;
//        delete m_pgbTiltChkDM;
//        delete m_pgbEncJumpTHChkDM;
//        delete m_pgbEncJumpRChkDM;
//        delete m_pgbConductivityChkDM;
//    }
    // set Analyse Advice group
    m_pgbAnalyseAdvDM = new QButtonGroup( this);
    m_pgbAnalyseAdvDM->addButton(ui->rbNFFDM, 0);
    m_pgbAnalyseAdvDM->addButton(ui->rbWarrentyDM, 1);
    m_pgbAnalyseAdvDM->addButton(ui->rbGoodWillDM, 2);
    m_pgbAnalyseAdvDM->addButton(ui->rbWithCostsDM, 3);
    m_pgbAnalyseAdvDM->addButton(ui->rbScrapItemDM, 4);
    //ui->rbNFFDM->setChecked(true);

    // set Causer group
    m_pgbCauserDM = new QButtonGroup( this);
    m_pgbCauserDM->addButton( ui->rbCustomerDM, 0);
    m_pgbCauserDM->addButton( ui->rbASYSDM, 1);
    //ui->rbCustomerDM->setChecked( true);

    // set DM repair or not
    m_pgbRepairDMUpgradeChk = new QButtonGroup( this);
    m_pgbRepairDMUpgradeChk->addButton( ui->rbRepairDM, 1);
    m_pgbRepairDMUpgradeChk->addButton( ui->rbDMIsOK, 0);
    //ui->rbRepairDM->setChecked( true);

    // set DM Analyse result
//    m_pgbAnalyseChkDM = new QButtonGroup( this);
//    m_pgbAnalyseChkDM->addButton( ui->rbDMAnalyseOK, 0);
//    m_pgbAnalyseChkDM->addButton( ui->rbDMAnalyseNOK, 1);
//    ui->rbDMAnalyseNOK->setChecked( true);

    // set DM surface damage result
    m_pgbSurfaceDamegeChkDM = new QButtonGroup( this);
    m_pgbSurfaceDamegeChkDM->addButton( ui->rbSurfNotDamegeOKDM, 0);
    m_pgbSurfaceDamegeChkDM->addButton( ui->rbSurfNotDamegeNOKDM, 1);
    m_pgbSurfaceDamegeChkDM->addButton( ui->rbSurfNotDamegeNADM, 2);     //new code****************************
    //ui->rbSurfNotDamegeNOKDM->setChecked( true);

    // set DM Electricity result
    m_pgbEleChkDM = new QButtonGroup( this);
    m_pgbEleChkDM->addButton( ui->rbEletricityOKDM, 0);
    m_pgbEleChkDM->addButton( ui->rbEletricityNOKDM, 1);
    m_pgbEleChkDM->addButton( ui->rbEletricityNADM, 2);     //new code****************************
    //ui->rbEletricityNOKDM->setChecked( true);

    // set DM function result
    m_pgbFunChkDM = new QButtonGroup( this);
    m_pgbFunChkDM->addButton( ui->rbFunTestOKDM, 0);
    m_pgbFunChkDM->addButton( ui->rbFunTestNOKDM, 1);
    m_pgbFunChkDM->addButton( ui->rbFunTestNADM, 2);     //new code****************************
    //ui->rbFunTestNOKDM->setChecked( true);

    // set DM Data transfer result
    m_pgbDataTransChkDM = new QButtonGroup( this);
    m_pgbDataTransChkDM->addButton( ui->rbDataTransTestOKDM, 0);
    m_pgbDataTransChkDM->addButton( ui->rbDataTransTestNOKDM, 1);
    m_pgbDataTransChkDM->addButton( ui->rbDataTransTestNADM, 2);     //new code****************************
    //ui->rbDataTransTestNOKDM->setChecked( true);

    // set DM TH Motor
    m_pgbTHMotorChkDM = new QButtonGroup( this);
    m_pgbTHMotorChkDM->addButton( ui->rbTHMotorOK, 0);
    m_pgbTHMotorChkDM->addButton( ui->rbTHMotorNOK, 1);
    m_pgbTHMotorChkDM->addButton( ui->rbTHMotorNA, 2);     //new code****************************
    //ui->rbTHMotorNOK->setChecked( true);

    // set DM R Motor
    m_pgbRMotorChkDM = new QButtonGroup( this);
    m_pgbRMotorChkDM->addButton( ui->rbRMotorOK, 0);
    m_pgbRMotorChkDM->addButton( ui->rbRMotorNOK, 1);
    m_pgbRMotorChkDM->addButton( ui->rbRMotorNA, 2);     //new code****************************
    //ui->rbRMotorNOK->setChecked( true);

    // set DM TH Gear
    m_pgbTHGearChkDM = new QButtonGroup( this);
    m_pgbTHGearChkDM->addButton( ui->rbTHGearOK, 0);
    m_pgbTHGearChkDM->addButton( ui->rbTHGearNOK, 1);
    m_pgbTHGearChkDM->addButton( ui->rbTHGearNA, 2);     //new code****************************
    //ui->rbTHGearNOK->setChecked( true);

    // set DM R Gear
    m_pgbRGearChkDM = new QButtonGroup( this);
    m_pgbRGearChkDM->addButton( ui->rbRGearOK, 0);
    m_pgbRGearChkDM->addButton( ui->rbRGearNOK, 1);
    m_pgbRGearChkDM->addButton( ui->rbRGearNA, 2);     //new code****************************
    //ui->rbRGearNOK->setChecked( true);

    // set DM Tilt ok or not
    m_pgbTiltChkDM = new QButtonGroup( this);
    m_pgbTiltChkDM->addButton( ui->rbTiltOK, 0);
    m_pgbTiltChkDM->addButton( ui->rbTiltNOK, 1);
    m_pgbTiltChkDM->addButton( ui->rbTiltNA, 2);     //new code****************************
    //ui->rbTiltNOK->setChecked( true);

    // set DM Encoder Jump TH
    m_pgbEncJumpTHChkDM = new QButtonGroup( this);
    m_pgbEncJumpTHChkDM->addButton( ui->rbEncJumpTestTHOK, 0);
    m_pgbEncJumpTHChkDM->addButton( ui->rbEncJumpTestTHNOK, 1);
    m_pgbEncJumpTHChkDM->addButton( ui->rbEncJumpTestTHNA, 2);     //new code****************************
    //ui->rbEncJumpTestTHNOK->setChecked( true);

    // set DM Encoder Jump R
    m_pgbEncJumpRChkDM = new QButtonGroup( this);
    m_pgbEncJumpRChkDM->addButton( ui->rbEncJumpTestROK, 0);
    m_pgbEncJumpRChkDM->addButton( ui->rbEncJumpTestRNOK, 1);
    m_pgbEncJumpRChkDM->addButton( ui->rbEncJumpTestRNA, 2);     //new code****************************
    //ui->rbEncJumpTestRNOK->setChecked( true);

    // set DM conductivity
    m_pgbConductivityChkDM = new QButtonGroup( this);
    m_pgbConductivityChkDM->addButton( ui->rbConductivityChkOKDM, 0);
    m_pgbConductivityChkDM->addButton( ui->rbConductivityChkNOKDM, 1);
    m_pgbConductivityChkDM->addButton( ui->rbConductivityChkNADM, 2);     //new code****************************
    //ui->rbConductivityChkNOKDM->setChecked( true);
}

// for ZT
void CAnalyseData::setRadioButtonsIDsInGB4ZT( void)
{
    // set Analyse Advice group
    m_pgbAnalyseAdvZT = new QButtonGroup( this);
    m_pgbAnalyseAdvZT->addButton(ui->rbNFFZT, 0);
    m_pgbAnalyseAdvZT->addButton(ui->rbWarrentyZT, 1);
    m_pgbAnalyseAdvZT->addButton(ui->rbGoodWillZT, 2);
    m_pgbAnalyseAdvZT->addButton(ui->rbWithCostsZT, 3);
    m_pgbAnalyseAdvZT->addButton(ui->rbScrapItemZT, 4);
    //ui->rbNFFZT->setChecked(true);

    // set Causer group
    m_pgbCauserZT = new QButtonGroup( this);
    m_pgbCauserZT->addButton( ui->rbCustomerZT, 0);
    m_pgbCauserZT->addButton( ui->rbASYSZT, 1);
    //ui->rbCustomerZT->setChecked( true);

    // set ZT repair or not
    m_pgbRepairZTUpgradeChk = new QButtonGroup( this);
    m_pgbRepairZTUpgradeChk->addButton( ui->rbRepairZT, 1);
    m_pgbRepairZTUpgradeChk->addButton( ui->rbZTIsOK, 0);
    //ui->rbRepairZT->setChecked( true);

    // set ZT Analyse result
//    m_pgbAnalyseChkZT = new QButtonGroup( this);
//    m_pgbAnalyseChkZT->addButton( ui->rbZTAnalyseOK, 0);
//    m_pgbAnalyseChkZT->addButton( ui->rbZTAnalyseNOK, 1);
//    ui->rbZTAnalyseNOK->setChecked( true);

    // set ZT surface damage result
    m_pgbSurfaceDamegeChkZT = new QButtonGroup( this);
    m_pgbSurfaceDamegeChkZT->addButton( ui->rbSurfNotDamegeOKZT, 0);
    m_pgbSurfaceDamegeChkZT->addButton( ui->rbSurfNotDamegeNOKZT, 1);
    m_pgbSurfaceDamegeChkZT->addButton( ui->rbSurfNotDamegeNAZT, 2);    //new code**************************
    //ui->rbSurfNotDamegeNOKZT->setChecked( true);

    // set ZT LVDT
    m_pgbLVDTZT = new QButtonGroup( this);
    m_pgbLVDTZT->addButton( ui->rbLVDTOK, 0);
    m_pgbLVDTZT->addButton( ui->rbLVDTNOK, 1);
    m_pgbLVDTZT->addButton( ui->rbLVDTNA, 2);       //new code**************************
    //ui->rbLVDTNOK->setChecked( true);

    // set ZT Ref Sensor
    m_pgbRefSenZT = new QButtonGroup( this);
    m_pgbRefSenZT->addButton( ui->rbRefSensorOK, 0);
    m_pgbRefSenZT->addButton( ui->rbRefSensorNOK, 1);
    m_pgbRefSenZT->addButton( ui->rbRefSensorNA, 2);       //new code**************************
    //ui->rbRefSensorNOK->setChecked( true);

    // set ZT Z Motor
    m_pgbZMotorZT = new QButtonGroup( this);
    m_pgbZMotorZT->addButton( ui->rbZMotorOK, 0);
    m_pgbZMotorZT->addButton( ui->rbZMotorNOK, 1);
    m_pgbZMotorZT->addButton( ui->rbZMotorNA, 2);       //new code**************************
   // ui->rbZMotorNOK->setChecked( true);

    // set ZT Running Noise
    m_pgbRunningNoiseZT = new QButtonGroup( this);
    m_pgbRunningNoiseZT->addButton( ui->rbRunNoiseOK, 0);
    m_pgbRunningNoiseZT->addButton( ui->rbRunNoiseNOK, 1);
    m_pgbRunningNoiseZT->addButton( ui->rbRunNoiseNA, 2);       //new code**************************
    //ui->rbRunNoiseNOK->setChecked( true);

    // set ZT Cable
    m_pgbCableZT = new QButtonGroup( this);
    m_pgbCableZT->addButton( ui->rbCableOK, 0);
    m_pgbCableZT->addButton( ui->rbCableNOK, 1);
    m_pgbCableZT->addButton( ui->rbCableNA, 2);       //new code**************************
    //ui->rbCableNOK->setChecked( true);

    // set ZT conductivity
    m_pgbConductivityChkZT = new QButtonGroup( this);
    m_pgbConductivityChkZT->addButton( ui->rbConductivityChkOKZT, 0);
    m_pgbConductivityChkZT->addButton( ui->rbConductivityChkNOKZT, 1);
    m_pgbConductivityChkZT->addButton( ui->rbConductivityChkNAZT, 2);   //new code**************************
    //ui->rbConductivityChkNOKZT->setChecked( true);
}

// ------------ WRITE EXCEL ------------------
void CAnalyseData::createAnalyseSheet()
{
//    // save the protocol to excel file
//    //m_filePathExcelTmp = "D:\\ASYS\\Projects\\Analyse_ASYS\\AnalyseTmp-Ray.xls";//QApplication::applicationDirPath()
//    #ifdef USEAPPLICATIONPATH //ANALYSETILT_0424
//      m_filePathExcelTmp = QApplication::applicationDirPath()+"\\ANALYSETILT_v1.0.xls";
//    #else
//      //m_filePathExcelTmp = "D:\\ASYS\\Projects\\Analyse_ASYS\\AnalyseTmp-Ray.xls";
//      m_filePathExcelTmp = "D:\\ASYS\\Projects\\Analyse_ASYS\\ANALYSETILT_0728.xls";
//    #endif
    ui->lbStatus->setText("Creating Analyse Sheet");
    progressSave(3);
    closeExcel();
    //closeExcel1();    
//    on_pbSaveIni_clicked();
    progressSave(6);
//    createLabelFile();
    progressSave(10);
    buildProtocolTable();
//    qDebug() << m_filePathExcel;
    if(QFile::exists(m_filePathExcel))
    {
        QFile::remove(m_filePathExcel);
    }

    m_objExcel = new QAxObject("Excel.Application");
    if( m_objExcel==nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Excel is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }

    m_objExcel->setProperty("Visible",false);
    QAxObject* workbooks = m_objExcel->querySubObject("WorkBooks");// get the workbook.
    if( workbooks==nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Office is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }
    workbooks->setProperty("Visible",false);    //new code*********************
    workbooks->dynamicCall("Open (const QString&)", m_filePathExcelTmp);   // Open the file；
    m_objWorkbook = m_objExcel->querySubObject("ActiveWorkBook"); // Get the active workbook.
    progressSave(20);
    int item_count = 1;// for progress
    foreach(sPROTOCOLITEM item, vecProtocolItems)
    {
        if(item.p_func)
        {
            (this->*(item.p_func))(m_objWorkbook, item);
        }
        else
        {
            writeGeneralItem(m_objWorkbook, item);
        }
        progressSave(20 + (item_count++)*70/vecProtocolItems.size());
    }

    //-------------------------- For creating DataMatrix -----------------------------------// *********************************
    QAxObject *worksheet = m_objWorkbook->querySubObject("Sheets(int)", 4); // get 4th sheet ("Label" sheet)
    QSettings settings(m_settingFilePath,QSettings::IniFormat);
    settings.beginGroup("PathSetting");
    //QString strDMFilePath = settings.value("DataMatrixFilePath").toString();
    QString strDMFolder = QDir::toNativeSeparators(settings.value("DataMatrixFilePath").toString());

    QAxObject* rangeA1 = worksheet->querySubObject("Range())", "A1");       //ex: The cell A1
    QString strValueA1 = rangeA1->dynamicCall("Value()").toString();        //strRange: the raw value of cell A1

    QAxObject* rangeA2 = worksheet->querySubObject("Range())", "A2");       //ex: The cell A2
    QString strValueA2 = rangeA2->dynamicCall("Value()").toString();        //strRange: the raw value of cell A2

    QAxObject* rangeA3 = worksheet->querySubObject("Range())", "A3");       //ex: The cell A3
    QString strValueA3 = rangeA3->dynamicCall("Value()").toString();        //strRange: the raw value of cell A3

    //                               strText,    strSavePath,   strFileName,                   iFileType, iSize //1: jpg, 2: jpeg, 3. png
    libDataMatrix.GenerateDataMatrix(strValueA1, strDMFolder, (m_fileName + "_DataMatrix_ARM"),   1,        200);    //(m_fileName + "_QrCode_ARM")
    libDataMatrix.GenerateDataMatrix(strValueA2, strDMFolder, (m_fileName + "_DataMatrix_DM"),    1,        200);
    libDataMatrix.GenerateDataMatrix(strValueA3, strDMFolder, (m_fileName + "_DataMatrix_Z"),     1,        200);

    // 18, 17, 17 -> 0.25 inch
    insertDataMatrix(m_objWorkbook, 4, "A1", (strDMFolder + "\\" + m_fileName + "_DataMatrix_ARM.jpg"), 21.6);
    insertDataMatrix(m_objWorkbook, 4, "A2", (strDMFolder + "\\" + m_fileName + "_DataMatrix_DM.jpg"), 22);
    insertDataMatrix(m_objWorkbook, 4, "A3", (strDMFolder + "\\" + m_fileName + "_DataMatrix_Z.jpg"), 22);
//    //-----------------------------------------------------------------------------------------------------------------//

    // add to disable checking compatibility
    m_objWorkbook->setProperty("DisplayAlerts", false);
    m_objWorkbook->setProperty("CheckCompatibility", false);
    m_objWorkbook->setProperty("DoNotPromptForConvert", true);

    m_objWorkbook->dynamicCall("SaveAs(const QString&)",
                                     QDir::toNativeSeparators(m_filePathExcel));

    progressSave(98);

    closeExcel();
    progressSave(100);
    qDebug() << "Analyse Sheet created.";
    ui->lbStatus->setText("Analyse Sheet was created");
}

void CAnalyseData::closeExcel( void)
{
    if(m_objWorkbook != nullptr)
    {
        m_objWorkbook->dynamicCall("Close (Boolean)", false);
        m_objWorkbook = nullptr;
    }// if()

    if(m_objExcel != nullptr)
    {
        m_objExcel->dynamicCall("Quit (void)");
        delete m_objExcel;
        m_objExcel = nullptr;
    }// if()
    vecProtocolItems.clear();
    vecRepairItems.clear();

//    if(m_objLabelExcel != nullptr)
//    {
//       m_objLabelExcel->dynamicCall("Quit (void)");
//       delete m_objLabelExcel;
//       m_objLabelExcel = nullptr;
//    }

//    if(m_objLabelWorkbook != nullptr)
//    {
//       m_objLabelWorkbook->dynamicCall("Quit (void)");
//       delete m_objLabelWorkbook;
//       m_objLabelWorkbook = nullptr;
//    }
}

void CAnalyseData::writeRobotTypeItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0, iRobType = 0;
    getRowColumn(item.Cell, &row, &column);
    iRobType = item.Value.toInt();// row++
    //qDebug()<<"iRobType = "<< iRobType<<"Id = "<< m_pgbRobotTypeARM->checkedId()<<"Btn = "<< m_pgbRobotTypeARM->checkedButton()<< endl;
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row+iRobType, column);
    QString value = "R";
    switch(iRobType)
    {
        case Robot_AAR:
          value.append(" AAR:");
        break;
        case Robot_AARNT:
          value.append(" AARNT:");
        break;
        case Robot_NXT:
          value.append(" NXT:");
        break;
    }
    cell->setProperty("Value", value);  // Set the value in the cell.

//    QAxObject *font = cell->querySubObject("Font");
//    font->setProperty("Size", 72);
//    font->setProperty("Bold", false);
//    font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//    font->setProperty("Color", 0);

    QAxObject* chars1 = cell->querySubObject("Characters(int, int)", 0, 1);
    QAxObject *font1 = chars1->querySubObject("Font");
//    font1->setProperty("Size", 72);
    font1->setProperty("Name", QStringLiteral("Wingdings 2"));
}

void CAnalyseData::writeAdviceItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0, iAdvice = 0;
    getRowColumn(item.Cell, &row, &column);
    iAdvice = item.Value.toInt();// row++

    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row+iAdvice, column);
    QString value = "U";//"R";
    switch(iAdvice)
    {
        case NFF:
          value.append(" NFF");
        break;
        case Warrenty:
          value.append(" Warrenty");
        break;
        case GoodWill:
          value.append(" Good Will");
        break;
        case WithCosts:
          value.append(" with costs");
        break;
        case ScrapItem:
          value.append(" Scrap Item");
        break;
    }
    cell->setProperty("Value", value);  // Set the value in the cell.

//    QAxObject *font = cell->querySubObject("Font");
//    font->setProperty("Size", 12);
//    font->setProperty("Bold", false);
//    font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//    font->setProperty("Color", 0);

    QAxObject* chars1 = cell->querySubObject("Characters(int, int)", 0, 1);
    QAxObject *font1 = chars1->querySubObject("Font");
    font1->setProperty("Size", 12);
    font1->setProperty("Name", QStringLiteral("Wingdings 2"));
}

void CAnalyseData::writeCauserItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0, iCauser = 0;
    getRowColumn(item.Cell, &row, &column);
    iCauser = item.Value.toInt();// row++

    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row+iCauser, column);
    QString value = "U";//"R";
    switch(iCauser)
    {
        case Customer:
          value.append(" Customer");
        break;
        case ASYS:
          value.append(" ASYS");
        break;
    }
    cell->setProperty("Value", value);  // Set the value in the cell.

//    QAxObject *font = cell->querySubObject("Font");
//    font->setProperty("Size", 72);
//    font->setProperty("Bold", false);
//    font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//    font->setProperty("Color", 0);

    QAxObject* chars1 = cell->querySubObject("Characters(int, int)", 0, 1);
    QAxObject *font1 = chars1->querySubObject("Font");
    font1->setProperty("Size", 12);
    font1->setProperty("Name", QStringLiteral("Wingdings 2"));//Wingdings 2
}

void CAnalyseData::writeGeneralItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0;
    getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
//    QString strCells[JUDGEITEMS] = {"D14", "D15", "D16", "D17", "D18", "D19", "D20", "D21", "D27", "D28"};

    switch(item.Value.type())
    {
        case QVariant::Int:
        {
            QString value = "R";
            int iRst = item.Value.toInt();
            //qDebug()<<"iRst = "<<iRst<<endl;
            QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column+iRst);
            QAxObject *font = cell->querySubObject("Font");
            font->setProperty("Size", 12);

            switch(iRst)
            {
                case Test_OK:
                  value.append(" OK");
                  font->setProperty("Color", QColor(Qt::black));    //******************
                break;
                case Test_NG:
                  value.append(" NOK");
                  font->setProperty("Color", QColor(Qt::red));    //******************
                break;
                case Test_NA:
                  value.append(" N/A");
                  font->setProperty("Color", QColor(Qt::black));    //******************
                break;
            }
            cell->setProperty("Value", value);

//            font->setProperty("Bold", false);
//            font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//            font->setProperty("Color", 0); // WdColorBlack = 0

            QAxObject* chars1 = cell->querySubObject("Characters(int, int)", 0, 1);
            QAxObject *font1 = chars1->querySubObject("Font");
            font1->setProperty("Size", 12);
            font1->setProperty("Name", QStringLiteral("Wingdings 2"));//Wingdings2
            break;
        }
        case QVariant::String:
        {
            QString value = item.Value.toString();
            cell->setProperty("Value", value);  // Set the value in the cell.

//            QAxObject *font = cell->querySubObject("Font");
//            font->setProperty("Size", 72);
//            font->setProperty("Bold", false);
//            font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//            font->setProperty("Color", 0); // WdColorBlack = 0
            break;
        }
        default:
        break;
    }// switch()
}

void CAnalyseData::writeRobotType(QAxObject *workbook, CAnalyseData::sPROTOCOLITEM item)    //new*******************
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0;
    getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    QString strRobotType = item.Value.toString();
    QString strValue = "";
    if(strRobotType.contains("NT"))
    {
        strValue = "NT";
    }else if(strRobotType.contains("NXT"))
    {
        strValue = "NXT";
    }else if(strRobotType.contains("DF"))
    {
        strValue = "DF";
    }
    else
    {
        strValue = "SC";
    }

    cell->setProperty("Value", strValue);  // Set the value in the cell.
}

void CAnalyseData::writeCurrentItem(QAxObject *workbook, CAnalyseData::sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0;
    getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    QString strNumValue = item.Value.toString();
    QString strValue = strNumValue + " A";
    cell->setProperty("Value", strValue);  // Set the value in the cell.
}

void CAnalyseData::writeZTSNFormerItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0;
    getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    QString value = "ARE-";
    QString lenVal = "";
    switch(item.Value.toInt())
    {
        case Len35mm:
          lenVal = "035";
        break;
        case Len50mm:
          lenVal = "050";
        break;
        case Len65mm:
          lenVal = "065";
        break;
    }

    value += lenVal + "-XX-XX-";
    cell->setProperty("Value", value);  // Set the value in the cell.

//    QAxObject *font = cell->querySubObject("Font");
//    font->setProperty("Size", 72);
//    font->setProperty("Bold", false);
//    font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//    font->setProperty("Color", 0); // WdColorBlack = 0
}

void CAnalyseData::writeRepairOrNotItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0, iRepairOrNot = 0;
    getRowColumn(item.Cell, &row, &column);
    iRepairOrNot = item.Value.toInt();


    QString value = "R";
    switch(iRepairOrNot)
    {
        case Test_OK:
          iRepairOrNot = 2;
        break;
        case Test_NG:
          iRepairOrNot = 0;
        break;
    }
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column+iRepairOrNot);
    cell->setProperty("Value", value);  // Set the value in the cell.

//    QAxObject *font = cell->querySubObject("Font");
//    font->setProperty("Size", 72);
//    font->setProperty("Bold", false);
//    font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//    font->setProperty("Color", 0);

    QAxObject* chars1 = cell->querySubObject("Characters(int, int)", 0, 1);
    QAxObject *font1 = chars1->querySubObject("Font");
//    font1->setProperty("Size", 72);
    font1->setProperty("Name", QStringLiteral("Wingdings 2"));
}

void CAnalyseData::writeUpgradeItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0, iUpgrade = 0;
    getRowColumn(item.Cell, &row, &column);
    iUpgrade = item.Value.toInt();

    QString value = "R";
    if(iUpgrade==Test_NG)
    {
        QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
        cell->setProperty("Value", value);  // Set the value in the cell.

//        QAxObject *font = cell->querySubObject("Font");
//        font->setProperty("Size", 72);
//        font->setProperty("Bold", false);
//        font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//        font->setProperty("Color", 0);

        QAxObject* chars1 = cell->querySubObject("Characters(int, int)", 0, 1);
        QAxObject *font1 = chars1->querySubObject("Font");
//        font1->setProperty("Size", 72);
        font1->setProperty("Name", QStringLiteral("Wingdings 2"));
    }// if()

}

void CAnalyseData::writeCommentsItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0;
    getRowColumn(item.Cell, &row, &column);
   // qDebug()<<"Gen robType= "<<robType<<",Row = "<<row<<",col = "<<column<<endl;
    //QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    //QString value = "";
    QStringList commentsList = item.Value.toString().split("\n");

    for( int i = 0; i < commentsList.count(); i++)
    {
       if(commentsList[i]!="")
       {
//           qDebug()<<"comments= "<<commentsList[i]<<endl;
           QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row+i, column);
           cell->setProperty("Value", commentsList[i]);

//           QAxObject *font = cell->querySubObject("Font");
//           font->setProperty("Size", 72);
//           font->setProperty("Bold", false);
//           font->setProperty("Name", QStringLiteral("Czcionka tekstu podstawowego"));
//           font->setProperty("Color", 0); // WdColorBlack = 0
       }// if()
    }// for()
}

void CAnalyseData::insertImgItem(QAxObject* workbook, sPROTOCOLITEM item)
{// B45>>>B36
    //qDebug()<<"imgPaht = "<<m_imgPath<<endl;
    //m_imgPath = "D:/ASYS/Projects/Analyse_ASYS/protocols/NT2685.JPG";
    //"D://ASYS//Projects//Analyse_ASYS//protocols//NT2685.JPG";>>NG
    // D:/ASYS/Projects/Analyse_ASYS/protocols/NT2685.JPG>>>NG
    // D:\ASYS\Projects\Analyse_ASYS\protocols\NT2685.JPG">>>OK
    // D:\\ASYS\\Projects\Analyse_ASYS\\protocols\\NT2685.JPG>>>OK

    //m_imgPath = "D:\\ASYS\\Projects\Analyse_ASYS\\protocols\\NT2685.JPG";
//    qDebug()<<"imgPaht = "<<m_imgPath<<endl;
    m_imgPath = QDir::toNativeSeparators(m_imgPath);
    //qDebug()<<"imgPahtt = "<<m_imgPath<<endl;
    QString valTmp = ui->lbImageName->text();
    if(QFile::exists(m_imgPath)&& valTmp!="")
    {
        //
        //qDebug()<<"img = "<<m_imgPath<<endl;
        QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.

        QAxObject* range = worksheet->querySubObject("Range(const QString&)", item.Cell);
        QImage image(m_imgPath);

        // Resize image to desired size
        QSize size(ImageResizeWidth, ImageResizeHeight);
        image = image.scaled(size, Qt::KeepAspectRatio);
        // Convert image to pixmap
        //QPixmap pixmap = QPixmap::fromImage(image);

        QAxObject* pictures = worksheet->querySubObject("Pictures()");
        QAxObject* picture = pictures->querySubObject("Insert(const QString&, bool)", m_imgPath, true);

        int width = image.width();
        int height = image.height();
        picture->setProperty("Width", width);
        picture->setProperty("Height", height);//height/2

        double left = range->property("Left").toDouble();
        double top = range->property("Top").toDouble();

        picture->setProperty("Left", left);
        picture->setProperty("Top", top);
    }
    // 20230706 add
    else
    {// get BID code
        //
        QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
        int row = 0, column = 0;
        getRowColumn(item.Cell, &row, &column);
        QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
        QString strBID = "BID: ";
        QString strBIDVal = "";
        QString val = "";

        strBIDVal = ui->leBIDCode->text();

        val = strBID + strBIDVal;
        cell->setProperty("Value", val);


    }
}

void CAnalyseData::insertRemarkImg(QAxObject* workbook, sPROTOCOLITEM item)
{
    QString strImgPath = item.Value.toString();

    strImgPath = QDir::toNativeSeparators(strImgPath);

    if(strImgPath != "")
    {
        if(QFile::exists(strImgPath))
        {
            QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.

            QAxObject* range = worksheet->querySubObject("Range(const QString&)", item.Cell);
            QImage image(strImgPath);
            // Resize image to desired size
            QSize size(ImageResizeWidth, ImageResizeHeight);
            image = image.scaled(size, Qt::KeepAspectRatio);
            // Convert image to pixmap
            //QPixmap pixmap = QPixmap::fromImage(image);

            int width = image.width();
            int height = image.height();

            QAxObject* pictures = worksheet->querySubObject("Pictures()");
            QAxObject* picture = pictures->querySubObject("Insert(const QString&, bool)", strImgPath, true);

            picture->setProperty("Width", width);
            picture->setProperty("Height", height);//height/2

            double left = range->property("Left").toDouble();
            double top = range->property("Top").toDouble();

            picture->setProperty("Left", left);
            picture->setProperty("Top", top);
        }
    }
}

void CAnalyseData::chkTextCokorItem(QAxObject* workbook, sPROTOCOLITEM item)
{
    QString strArmCells[10] = {"D14", "D15", "D16", "D17", "D18", "D19", "D20", "D21", "D27", "D28"};
    QString strDMCells[10] = {"D14", "D15", "D17", "D18", "D24", "D25", "D32", "D33", "F17", "F18"};
    QString strZCells[8] = {"F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21"};
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", item.Sheet);   // Get the worksheet.
    int row = 0, column = 0;
    getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    QString value = item.Value.toString();
    cell->setProperty("Value", value);  // Set the value in the cell.
    QAxObject *font = cell->querySubObject("Font");
    font->setProperty("Color", 0);
    font->setProperty("Size", 12);
    if(item.Sheet==AnalyseARM)
    {
        for(int i=0; i<10; i++ )
        {
            if(strArmCells[i]==item.Cell && m_bChkColor[i])
            {
                font->setProperty("Color", QColor(Qt::red));
            }
        }
    }
    else if(item.Sheet==AnalyseDM)
    {
        for(int j=0; j<=7; j++ )
        {
            if(strDMCells[j]==item.Cell&&m_bChkColor[j+10])
            {
                font->setProperty("Color", QColor(Qt::red));
            }
        }
        if(strDMCells[8]==item.Cell&&m_bChkColor[12])
        {
            font->setProperty("Color", QColor(Qt::red));
        }
        if(strDMCells[9]==item.Cell&&m_bChkColor[13])
        {
            font->setProperty("Color", QColor(Qt::red));
        }
    }
    else if(item.Sheet==AnalyseZT)
    {
        cell->setProperty("HorizontalAlignment", -4108);
        for(int k=0; k<=8; k++)
        {
            if(strZCells[k]==item.Cell && m_bChkColor[k+18])
            {
                font->setProperty("Color", QColor(Qt::red));
            }
        }
    }
}

void CAnalyseData::getRowColumn(QString cell, int* row, int* column)
{
    QRegExp rx("([A-Z]+)(\\d+)");
    rx.indexIn(cell);

    QByteArray bytes = rx.cap(1).toLocal8Bit();
    for(int i = 0; i < bytes.size(); i++)
    {
        *column += bytes[i] - 'A' + 1;
    }
    *row = rx.cap(2).toInt();
}

void CAnalyseData::progressSave(int value)
{
    ui->pgbProcess->setValue(value);
}

// ------------ BUTTONS/TEXT -------------------
void CAnalyseData::on_pbImageName_clicked()
{
    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::ExistingFile);
    dialog.setOption(QFileDialog::DontUseNativeDialog);
    //dialog.setDirectory("//tw-srv-02/Data/06_SoftwareDev/03_Project/09_2023/02_2023_Digitized_AnalysisProtocol/06_Parameters");//m_imgDefaultPath = //tw-srv-02/Data/06_SoftwareDev/03_Project/09_2023/02_2023_Digitized_AnalysisProtocol/06_Parameters
    dialog.setDirectory(m_imgDefaultPath);//m_imgDefaultPath
    m_fileFrom= QFileDialog::getOpenFileName(this);// get file full path
    m_fileInfoFrom = QFileInfo(m_fileFrom);
//    qDebug()<<"fileInfo="<<m_fileInfoFrom.absoluteFilePath()<<endl;
    m_imgPath = m_fileInfoFrom.absoluteFilePath();//m_fileFrom;//
//    qDebug()<<"m_imgPath1="<<m_imgPath<<endl;
    m_fileNameFrom = m_fileInfoFrom.fileName();  // get file name
    m_filePathFrom = m_fileInfoFrom.absolutePath();// get file path without file name
    ui->lbImagePath->setText(m_filePathFrom);
//    qDebug() << m_filePathFrom;
    ui->lbImageName->setText(m_fileNameFrom);
    ui->leBIDCode->clear();// 20230706 add
    //ui->lbImageName->setText(m_imgPath);// show whole path for the image file.
}

void CAnalyseData::on_leRepairNRARM_textChanged(const QString &arg1)
{
    QString strTmp;
    m_strTmpTest = arg1;
    strTmp = ui->leRepairNRARM->text();
    ui->leRepairNRDM->setText(strTmp);
    ui->leRepairNRZT->setText(strTmp);
}

void CAnalyseData::on_leVacuumARM_textChanged(const QString &arg1)
{
    if(ui->leVacuumARM->text() == "N/A")
    {
        ui->leVacuumARM->setStyleSheet("color:red;");
        m_bChkColor[ARMVaccum] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leVacuumARM->text().toFloat();
        if(fVal > ARMVACUUM_LIMIT)
        {
            ui->leVacuumARM->setStyleSheet("color:red;");
            m_bChkColor[ARMVaccum] = true;
        }
        else
        {
            ui->leVacuumARM->setStyleSheet("color:black;");
            m_bChkColor[ARMVaccum] = false;
        }
    }
}

void CAnalyseData::on_leFlowARM_textChanged(const QString &arg1)
{
    if(ui->leFlowARM->text() == "N/A")
    {
        ui->leFlowARM->setStyleSheet("color:red;");
        m_bChkColor[ARMFlow] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leFlowARM->text().toFloat();

        if((fVal - TINYVALUE) <= ARMFLOW_UP && (fVal + TINYVALUE) >= ARMFLOW_DN)
        {
            ui->leFlowARM->setStyleSheet("color:black;");
            m_bChkColor[ARMFlow] = false;
        }
        else
        {
            ui->leFlowARM->setStyleSheet("color:red;");
            m_bChkColor[ARMFlow] = true;
        }
    }
}

void CAnalyseData::on_leOAG_textChanged(const QString &arg1)
{
    if(ui->leOAG->text() == "N/A")
    {
        ui->leOAG->setStyleSheet("color:red;");
        m_bChkColor[ARMOAG] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leOAG->text().toFloat();

        if(fVal <= OAG_UP && fVal >= OAG_DN)
        {
            ui->leOAG->setStyleSheet("color:black;");
            m_bChkColor[ARMOAG] = false;
        }
        else
        {
            ui->leOAG->setStyleSheet("color:red;");
            m_bChkColor[ARMOAG] = true;
        }
    }
}

void CAnalyseData::on_leUAG_textChanged(const QString &arg1)
{
    if(ui->leUAG->text() == "N/A")
    {
        ui->leUAG->setStyleSheet("color:red;");
        m_bChkColor[ARMUAG] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leUAG->text().toFloat();

        if((fVal - TINYVALUE) <= UAG_UP && (fVal + TINYVALUE) >= UAG_DN)
        {
            ui->leUAG->setStyleSheet("color:black;");
            m_bChkColor[ARMUAG] = false;
        }
        else
        {
            ui->leUAG->setStyleSheet("color:red;");
            m_bChkColor[ARMUAG] = true;
        }
    }
}

void CAnalyseData::on_leGeoRz_textChanged(const QString &arg1)
{
    if(ui->leGeoRz->text() == "N/A" || isDigitStr(ui->leGeoRz->text()) == false)
    {
        ui->leGeoRz->setStyleSheet("color:red;");
        m_bChkColor[ARMRz] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leGeoRz->text().toFloat();

        if((fVal - TINYVALUE) <= Rz_UP && (fVal + TINYVALUE) >= Rz_DN)
        {
            ui->leGeoRz->setStyleSheet("color:black;");
            m_bChkColor[ARMRz] = false;
        }
        else
        {
            ui->leGeoRz->setStyleSheet("color:red;");
            m_bChkColor[ARMRz] = true;
        }
    }
}

void CAnalyseData::on_leGeoRx_textChanged(const QString &arg1)
{
    if(ui->leGeoRx->text() == "N/A" || isDigitStr(ui->leGeoRx->text()) == false)
    {
        ui->leGeoRx->setStyleSheet("color:red;");
        m_bChkColor[ARMRx] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leGeoRx->text().toFloat();

        if((fVal - TINYVALUE) <= Rx_UP && (fVal + TINYVALUE) >= Rx_DN)
        {
            ui->leGeoRx->setStyleSheet("color:black;");
            m_bChkColor[ARMRx] = false;
        }
        else
        {
            ui->leGeoRx->setStyleSheet("color:red;");
            m_bChkColor[ARMRx] = true;
        }
    }
}

void CAnalyseData::on_leGeoRy_textChanged(const QString &arg1)
{
    if(ui->leGeoRy->text() == "N/A" || isDigitStr(ui->leGeoRy->text()) == false)
    {
        ui->leGeoRy->setStyleSheet("color:red;");
        m_bChkColor[ARMRy] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leGeoRy->text().toFloat();

        if((fVal - TINYVALUE) <= Ry_UP && (fVal + TINYVALUE) >= Ry_DN)
        {
            ui->leGeoRy->setStyleSheet("color:black;");
            m_bChkColor[ARMRy] = false;
        }
        else
        {
            ui->leGeoRy->setStyleSheet("color:red;");
            m_bChkColor[ARMRy] = true;
        }
    }
}

void CAnalyseData::on_leGeoDelHeight_textChanged(const QString &arg1)
{
    if(ui->leGeoDelHeight->text() == "N/A" || isDigitStr(ui->leGeoDelHeight->text()) == false)
    {
        ui->leGeoDelHeight->setStyleSheet("color:red;");
        m_bChkColor[ARMDeltaH] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leGeoDelHeight->text().toFloat();

        if((fVal - TINYVALUE) <= DeltaHeight_UP && (fVal + TINYVALUE) >= DeltaHeight_DN)
        {
            ui->leGeoDelHeight->setStyleSheet("color:black;");
            m_bChkColor[ARMDeltaH] = false;
        }
        else
        {
            ui->leGeoDelHeight->setStyleSheet("color:red;");
            m_bChkColor[ARMDeltaH] = true;
        }
    }
}

void CAnalyseData::on_le180DegVal_2_textChanged(const QString &arg1)
{
    if(ui->le180DegVal_2->text() == "N/A")
    {
        ui->le180DegVal->setText("N/A");
        ui->le180DegVal_2->setStyleSheet("color:red;");
        ui->le180DegVal->setStyleSheet("color:red;");
        m_bChkColor[DMMHDeg180] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->le180DegVal_2->text().toFloat();

        if((fVal - TINYVALUE) <= MicroHite_UP && (fVal + TINYVALUE) >= MicroHite_DN)
        {
            ui->le180DegVal_2->setStyleSheet("color:black;");
            ui->le180DegVal->setStyleSheet("color:black;");
            m_bChkColor[DMMHDeg180] = false;
        }
        else
        {
            ui->le180DegVal_2->setStyleSheet("color:red;");
            ui->le180DegVal->setStyleSheet("color:red;");
            m_bChkColor[DMMHDeg180] = true;
        }
    }
}

void CAnalyseData::on_le270DegVal_2_textChanged(const QString &arg1)
{
    if(ui->le270DegVal_2->text() == "N/A")
    {
        ui->le270DegVal->setText("N/A");
        ui->le270DegVal_2->setStyleSheet("color:red;");
        ui->le270DegVal->setStyleSheet("color:red;");
        m_bChkColor[DMMHDeg270] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->le270DegVal_2->text().toFloat();

        if((fVal - TINYVALUE) <= MicroHite_UP && (fVal + TINYVALUE) >= MicroHite_DN)
        {
            ui->le270DegVal_2->setStyleSheet("color:black;");
            ui->le270DegVal->setStyleSheet("color:black;");
            m_bChkColor[DMMHDeg270] = false;
        }
        else
        {
            ui->le270DegVal_2->setStyleSheet("color:red;");
            ui->le270DegVal->setStyleSheet("color:red;");
            m_bChkColor[DMMHDeg270] = true;
        }
    }
}

void CAnalyseData::on_leVacuumDM_textChanged(const QString &arg1)
{
    if(ui->leVacuumDM->text() == "N/A")
    {
        ui->leVacuumDM->setStyleSheet("color:red;");
        m_bChkColor[DMVaccum] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leVacuumDM->text().toFloat();
        if(fVal > ARMVACUUM_LIMIT)
        {
            ui->leVacuumDM->setStyleSheet("color:red;");
            m_bChkColor[DMVaccum] = true;
        }
        else if(ui->leVacuumDM->text() == "N/A")
        {
            ui->leVacuumDM->setStyleSheet("color:red;");
            m_bChkColor[DMVaccum] = true;
        }
        else
        {
            ui->leVacuumDM->setStyleSheet("color:black;");
            m_bChkColor[DMVaccum] = false;
        }
    }
}

void CAnalyseData::on_leFlowDM_textChanged(const QString &arg1)
{
    if(ui->leFlowDM->text() == "N/A")
    {
        ui->leFlowDM->setStyleSheet("color:red;");
        m_bChkColor[DMFlow] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leFlowDM->text().toFloat();

        if((fVal - TINYVALUE) <= ARMFLOW_UP && (fVal + TINYVALUE) >= ARMFLOW_DN)
        {
            ui->leFlowDM->setStyleSheet("color:black;");
            m_bChkColor[DMFlow] = false;
        }
        else
        {
            ui->leFlowDM->setStyleSheet("color:red;");
            m_bChkColor[DMFlow] = true;
        }
    }
}

void CAnalyseData::on_leRepPosPAR_textChanged(const QString &arg1)
{
    if(ui->leRepPosPAR->text() == "N/A")
    {
        ui->leRepPosPAR->setStyleSheet("color:red;");
        m_bChkColor[ARMPAR] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leRepPosPAR->text().toFloat();

        if((fVal - TINYVALUE) <= REPPOSPA_UP && (fVal + TINYVALUE) >= REPPOSPA_DN)
        {
            ui->leRepPosPAR->setStyleSheet("color:black;");
            m_bChkColor[ARMPAR] = false;
        }
        else
        {
            ui->leRepPosPAR->setStyleSheet("color:red;");
            m_bChkColor[ARMPAR] = true;
        }
    }
}

void CAnalyseData::on_leRepPosPATH_textChanged(const QString &arg1)
{
    if(ui->leRepPosPATH->text() == "N/A")
    {
        ui->leRepPosPATH->setStyleSheet("color:red;");
        m_bChkColor[ARMPATH] = true;
    }else
    {
        float fVal;
        m_strTmpTest = arg1;
        fVal = ui->leRepPosPATH->text().toFloat();

        if((fVal - TINYVALUE) <= REPPOSPA_UP && (fVal + TINYVALUE) >= REPPOSPA_DN)
        {
            ui->leRepPosPATH->setStyleSheet("color:black;");
            m_bChkColor[ARMPATH] = false;
        }
        else
        {
            ui->leRepPosPATH->setStyleSheet("color:red;");
            m_bChkColor[ARMPATH] = true;
        }
    }
}

void CAnalyseData::on_leCommutationTH_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leCommutationTH->setStyleSheet("color:red;");
        m_bChkColor[DMComTH] = true;
    }
}

void CAnalyseData::on_leCommutationR_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leCommutationR->setStyleSheet("color:red;");
        m_bChkColor[DMCOMR] = true;
    }
}

void CAnalyseData::on_leZeroingPosTH_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZeroingPosTH->setStyleSheet("color:red;");
        m_bChkColor[DM0PosTH] = true;
    }
}

void CAnalyseData::on_leZeroingPosR_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZeroingPosR->setStyleSheet("color:red;");
        m_bChkColor[DM0PosR] = true;
    }
}

void CAnalyseData::on_leZUpSCARA_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZUpSCARA->setStyleSheet("color:red;");
        m_bChkColor[ZSCCurrentUp] = true;
    }
}

void CAnalyseData::on_leZDownSCARA_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZDownSCARA->setStyleSheet("color:red;");
        m_bChkColor[ZSCCurrentDown] = true;
    }
}

void CAnalyseData::on_leZUpSCARANT_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZUpSCARANT->setStyleSheet("color:red;");
        m_bChkColor[ZNTCurrentUp] = true;
    }
}

void CAnalyseData::on_leZDownSCARANT_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZDownSCARANT->setStyleSheet("color:red;");
        m_bChkColor[ZNTCurrentDown] = true;
    }
}

void CAnalyseData::on_leZUpNXT_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZUpNXT->setStyleSheet("color:red;");
        m_bChkColor[ZNXTCurrentUp] = true;
    }
}

void CAnalyseData::on_leZDownNXT_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZDownNXT->setStyleSheet("color:red;");
        m_bChkColor[ZNXTCurrentDown] = true;
    }
}

void CAnalyseData::on_leZUpFA_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        qDebug() << "leZUpFA: " <<false;
        ui->leZUpFA->setStyleSheet("color:red;");
        m_bChkColor[ZFACurrentUp] = true;
    }
}

void CAnalyseData::on_leZDownFA_textChanged(const QString &arg1)
{
    if(arg1 == "N/A" || isDigitStr(arg1) == false)
    {
        ui->leZDownFA->setStyleSheet("color:red;");
        m_bChkColor[ZFACurrentDown] = true;
    }
}

//-------------------------------------------------------- Print Function --------------------------------------------------------//
void CAnalyseData::createPrintLabel()
{
    ui->lbStatus->setText("Creating Print Label");
    progressSave(30);
    createLabelFile();
    progressSave(60);

    //---- Print out label ----//
    QString filePath = mFilePathLabel;
//    qDebug() << "Print File: " << mFilePathLabel;
    QFile printFile(mFilePathLabel);
    if (!filePath.isEmpty() && printFile.exists())
    {
        progressSave(70);
        QAxObject* excel = new QAxObject("Excel.Application");
        excel->setProperty("Visible", false);
        QAxObject* workbooks = excel->querySubObject("Workbooks");
        QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", filePath);  // Open the file；
        QAxObject* worksheets = workbook->querySubObject("Worksheets");
        QAxObject* worksheet = worksheets->querySubObject("Item(int)", 1); //1th worksheet
        worksheet->setProperty("Visible", false);
        qDebug() << "Got print Workbook";
        progressSave(80);

        QAxObject* range = worksheet->querySubObject("UsedRange");
        QAxObject* font = range->querySubObject("Font");
        font->setProperty("Name", "Calibri"); // Font name
        font->setProperty("Size", 16.5);     // Font size in points
        font->setProperty("Bold", true);    // Bold
        progressSave(85);
        qDebug() << "Font Set";
        // Print the worksheet
        worksheet->dynamicCall("PrintOut()");

        workbook->setProperty("DisplayAlerts", false);
        workbook->setProperty("CheckCompatibility", false);
        workbook->setProperty("DoNotPromptForConvert", true);
        excel->setProperty("DisplayAlerts", false);
        excel->setProperty("CheckCompatibility", false);
        excel->setProperty("DoNotPromptForConvert", true);
        qDebug() << "Print Out";
        progressSave(90);
        workbook->dynamicCall("Close()");
        excel->dynamicCall("Quit()");
        qDebug() << "excel closed";
//        delete worksheet;   delete worksheets;  delete workbook;    delete workbooks;
//        delete excel;
        progressSave(100);
    }
    else
    {
       progressSave(0);
       QMessageBox::critical(NULL, "Error", "Print Error", QMessageBox::Yes, QMessageBox::Yes);
       return;
    }
    qDebug() << "Print Label created.";
    ui->lbStatus->setText("Print Label was created");

}

void CAnalyseData::on_pBtn_ImgRemark_ARM_clicked()
{
    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::ExistingFile);
    dialog.setOption(QFileDialog::DontUseNativeDialog);
    m_fileFrom= QFileDialog::getOpenFileName(this);// get file full path
    m_fileInfoFrom = QFileInfo(m_fileFrom);

    m_RemarkImgPath_ARM = m_fileInfoFrom.absoluteFilePath();//m_fileFrom;//
    m_RemarkImgFileNameARM = m_fileInfoFrom.fileName();  // get file name
    ui->lbl_ImgRemarkFileName_ARM->setText(m_RemarkImgFileNameARM);
}

void CAnalyseData::on_pBtn_ImgRemark_DM_clicked()
{
    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::ExistingFile);
    dialog.setOption(QFileDialog::DontUseNativeDialog);
    m_fileFrom= QFileDialog::getOpenFileName(this);// get file full path
    m_fileInfoFrom = QFileInfo(m_fileFrom);

    m_RemarkImgPath_DM = m_fileInfoFrom.absoluteFilePath();//m_fileFrom;//
    m_RemarkImgFileNameDM = m_fileInfoFrom.fileName();  // get file name
    ui->lbl_ImgRemarkFileName_DM->setText(m_RemarkImgFileNameDM);
}

void CAnalyseData::on_pBtn_ImgRemark_ZT_clicked()
{
    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::ExistingFile);
    dialog.setOption(QFileDialog::DontUseNativeDialog);
    m_fileFrom= QFileDialog::getOpenFileName(this);// get file full path
    m_fileInfoFrom = QFileInfo(m_fileFrom);

    m_RemarkImgPath_ZT = m_fileInfoFrom.absoluteFilePath();//m_fileFrom;//
    m_RemarkImgFileNameZT = m_fileInfoFrom.fileName();  // get file name
    ui->lbl_ImgRemarkFileName_ZT->setText(m_RemarkImgFileNameZT);
}

void CAnalyseData::insertDataMatrix(QAxObject *workbook, int iWorkSheet, QString strCell, QString strImgFileName, double dImgSize)
{
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", iWorkSheet);   // Get the worksheet.
    QAxObject* range = worksheet->querySubObject("Range(const QString&)", strCell);
    QAxObject* pictures = worksheet->querySubObject("Pictures()");
    QAxObject* picture = pictures->querySubObject("Insert(const QString&, bool)", strImgFileName, true);

    double dPic_XPosition = range->property("Left").toDouble() + range->property("Width").toDouble() - dImgSize - 1;
    double dPic_YPosition = range->property("Top").toDouble() + range->property("Height").toDouble() - dImgSize - 1;
    picture->setProperty("Width", dImgSize);
    picture->setProperty("Height", dImgSize);
    picture->setProperty("Left", dPic_XPosition);
    picture->setProperty("Top", dPic_YPosition);
}

bool CAnalyseData::isDigitStr(QString strSrc)
{
    QDoubleValidator validator;
    int pos = 0;
    QValidator::State result = validator.validate(strSrc, pos);

    if (result == QValidator::Acceptable || result == QValidator::Intermediate) // The userInput represents a valid double number
    {

        return true;
    }else //The userInput is not a valid double number
    {

        return false;
    }
}

void CAnalyseData::getLabelData()
{
    vecLabelData.resize(4);

    QString strRobSerial = m_robotType + ui->leRobotTypeSN->text();
    qDebug() << "Robot Serial: " << strRobSerial;

    vecLabelData[0] = strRobSerial;
    vecLabelData[1] = ui->leARMSN->text();

    QString strDMID = "";
    if(ui->leDMSN->text().contains("-"))
    {
        QStringList strLstDMSN = ui->leDMSN->text().split("-");
        if(strLstDMSN.size() >= 4)
            strDMID = strLstDMSN[3];
    }else
        strDMID = "R";

    QString strZID = "";
    if(ui->leZTSN2->text().contains("-"))
    {
        QStringList strLstZSN = ui->leZTSN2->text().split("-");
        if(strLstZSN.size() >= 4)
            strZID = strLstZSN[3];
    }else
        strZID = "R";

    vecLabelData[2] = strDMID;
    vecLabelData[3] = strZID;
}

void CAnalyseData::createLabelFile()
{
//    on_pbSaveIni_clicked();
    getLabelData();
    //--- Get file paths --//
    QSettings settings(m_settingFilePath,QSettings::IniFormat);
    settings.beginGroup("PathSetting");
    QString strLabelTemp = settings.value("TemplateVersion").toString() + "_Label";  //ANALYSETILT_v1.3_Label
    QString strLabelFolder = QDir::toNativeSeparators(settings.value("LabelFilePath").toString());
    QString strDMFolder = QDir::toNativeSeparators(settings.value("DataMatrixFilePath").toString());
    settings.endGroup();

    mFilePathLabel = strLabelFolder + "\\" + m_fileName + "_Label.xls";
    mFilePathLabelTemplate = QDir::toNativeSeparators(QApplication::applicationDirPath()) + "\\" + strLabelTemp + ".xls";

//    qDebug() << "Label Directory: " << mFilePathLabel;
//    qDebug() << "Label Template: " << mFilePathLabelTemplate;
//    qDebug() << "DM Folder: " << strDMFolder;

    //---- Create Label Excel File ----//
    closeExcel();
    if(QFile::exists(mFilePathLabel))
    {
        QFile::remove(mFilePathLabel);
    }

    m_objLabelExcel = new QAxObject("Excel.Application");
    if(m_objLabelExcel == nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Excel is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }

    m_objLabelExcel->setProperty("Visible",false);
    QAxObject* workbooks = m_objLabelExcel->querySubObject("WorkBooks");// get the workbook.
    if(workbooks == nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Office is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }
    workbooks->setProperty("Visible",false);    //new code*********************
    workbooks->dynamicCall("Open (const QString&)", mFilePathLabelTemplate);   // Open the file；
    m_objLabelWorkbook = m_objLabelExcel->querySubObject("ActiveWorkBook"); // Get the active workbook.

    //---- For inserting Label data ----//
    QAxObject *worksheetData = m_objLabelWorkbook->querySubObject("Sheets(int)", 2); // get 2th sheet ("Data" sheet)

    QAxObject* cellB1 = worksheetData->querySubObject("Cells(int,int)", 1, 2);
    cellB1->setProperty("Value", vecLabelData[0]);

    QAxObject* cellB2 = worksheetData->querySubObject("Cells(int,int)", 2, 2);
    cellB2->setProperty("Value", vecLabelData[1]);

    QAxObject* cellB3 = worksheetData->querySubObject("Cells(int,int)", 3, 2);
    cellB3->setProperty("Value", vecLabelData[2]);

    QAxObject* cellB4 = worksheetData->querySubObject("Cells(int,int)", 4, 2);
    cellB4->setProperty("Value", vecLabelData[3]);
//    qDebug() << "Label Data Written";

    //-------------------------- For creating DataMatrix -----------------------------------//
    QAxObject *worksheetLabel = m_objLabelWorkbook->querySubObject("Sheets(int)", 1); // get 1th sheet ("Label" sheet)

    QAxObject* rangeA1 = worksheetLabel->querySubObject("Range())", "A1");       //ex: The cell A1
    QString strValueA1 = rangeA1->dynamicCall("Value()").toString();        //strRange: the raw value of cell A1
    qDebug() << strValueA1;
    QAxObject* rangeA2 = worksheetLabel->querySubObject("Range())", "A2");       //ex: The cell A2
    QString strValueA2 = rangeA2->dynamicCall("Value()").toString();        //strRange: the raw value of cell A2
    qDebug() << strValueA2;
    QAxObject* rangeA3 = worksheetLabel->querySubObject("Range())", "A3");       //ex: The cell A3
    QString strValueA3 = rangeA3->dynamicCall("Value()").toString();        //strRange: the raw value of cell A3
    qDebug() << strValueA3;
    //                               strText,    strSavePath,   strFileName,                   iFileType, iSize //1: jpg, 2: jpeg, 3. png
    libDataMatrix.GenerateDataMatrix(strValueA1, strDMFolder, (m_fileName + "_DataMatrix_ARM"),   1,        200);    //(m_fileName + "_QrCode_ARM")
    libDataMatrix.GenerateDataMatrix(strValueA2, strDMFolder, (m_fileName + "_DataMatrix_DM"),    1,        200);
    libDataMatrix.GenerateDataMatrix(strValueA3, strDMFolder, (m_fileName + "_DataMatrix_Z"),     1,        200);

    //18, 17, 17 -> 0.25 inch
    insertDataMatrix(m_objLabelWorkbook, 1, "A1", (strDMFolder + "\\" + m_fileName + "_DataMatrix_ARM.jpg"), 21.6);
    insertDataMatrix(m_objLabelWorkbook, 1, "A2", (strDMFolder + "\\" + m_fileName + "_DataMatrix_DM.jpg"), 22);
    insertDataMatrix(m_objLabelWorkbook, 1, "A3", (strDMFolder + "\\" + m_fileName + "_DataMatrix_Z.jpg"), 22);
    qDebug() << "DataMatrix Inserted";
    //-----------------------------------------------------------------------------------------------------------------//

    // add to disable checking compatibility
    m_objLabelWorkbook->setProperty("DisplayAlerts", false);
    m_objLabelWorkbook->setProperty("CheckCompatibility", false);
    m_objLabelWorkbook->setProperty("DoNotPromptForConvert", true);

    m_objLabelWorkbook->dynamicCall("SaveAs(const QString&)",
                                     QDir::toNativeSeparators(mFilePathLabel));

//    qDebug() << "Label file Created";

    m_objLabelWorkbook->dynamicCall("Close()");
    m_objLabelExcel->dynamicCall("Quit()");
}



// ************** Create Repair Table ---------------------------------------------------------------------------- //


void CAnalyseData::setRadioButtonsIDsInGB4RepairARM( void){
    // set ArmBelts
    m_pgbRepairARM_ArmBelts = new QButtonGroup( this);
    m_pgbRepairARM_ArmBelts->addButton(ui->rbArmBeltsAvrPrice_OK, Repair_OK);
    m_pgbRepairARM_ArmBelts->addButton(ui->rbArmBeltsAvrPrice_Repair, Repair_Repair);
    m_pgbRepairARM_ArmBelts->addButton(ui->rbArmBeltsAvrPrice_NA, Repair_NA);

    // set UpperArmHousingUpgrade
    m_pgbRepairARM_UpperArmHousingUpgrade = new QButtonGroup( this);
    m_pgbRepairARM_UpperArmHousingUpgrade->addButton(ui->rbUpperArmHousingUpgrade_OK, Repair_OK);
    m_pgbRepairARM_UpperArmHousingUpgrade->addButton(ui->rbUpperArmHousingUpgrade_Repair, Repair_Repair);
    m_pgbRepairARM_UpperArmHousingUpgrade->addButton(ui->rbUpperArmHousingUpgrade_NA, Repair_NA);

    // set UpperArmHousing
    m_pgbRepairARM_UpperArmHousing = new QButtonGroup( this);
    m_pgbRepairARM_UpperArmHousing->addButton(ui->rbUpperArmHousing_OK, Repair_OK);
    m_pgbRepairARM_UpperArmHousing->addButton(ui->rbUpperArmHousing_Repair, Repair_Repair);
    m_pgbRepairARM_UpperArmHousing->addButton(ui->rbUpperArmHousing_NA, Repair_NA);

    // set UpperArmLid
    m_pgbRepairARM_UpperArmLid = new QButtonGroup( this);
    m_pgbRepairARM_UpperArmLid->addButton(ui->rbUpperArmLid_OK, Repair_OK);
    m_pgbRepairARM_UpperArmLid->addButton(ui->rbUpperArmLid_Repair, Repair_Repair);
    m_pgbRepairARM_UpperArmLid->addButton(ui->rbUpperArmLid_NA, Repair_NA);

    // set LowerArmHousingUpgrade
    m_pgbRepairARM_LowerArmHousingUpgrade = new QButtonGroup( this);
    m_pgbRepairARM_LowerArmHousingUpgrade->addButton(ui->rbLowerArmHousingUpgrade_OK, Repair_OK);
    m_pgbRepairARM_LowerArmHousingUpgrade->addButton(ui->rbLowerArmHousingUpgrade_Repair, Repair_Repair);
    m_pgbRepairARM_LowerArmHousingUpgrade->addButton(ui->rbLowerArmHousingUpgrade_NA, Repair_NA);

    // set LowerArmHousing
    m_pgbRepairARM_LowerArmHousing = new QButtonGroup( this);
    m_pgbRepairARM_LowerArmHousing->addButton(ui->rbLowerArmHousing_OK, Repair_OK);
    m_pgbRepairARM_LowerArmHousing->addButton(ui->rbLowerArmHousing_Repair, Repair_Repair);
    m_pgbRepairARM_LowerArmHousing->addButton(ui->rbLowerArmHousing_NA, Repair_NA);

    // set LowerArmLid
    m_pgbRepairARM_LowerArmLid = new QButtonGroup( this);
    m_pgbRepairARM_LowerArmLid->addButton(ui->rbLowerArmLid_OK, Repair_OK);
    m_pgbRepairARM_LowerArmLid->addButton(ui->rbLowerArmLid_Repair, Repair_Repair);
    m_pgbRepairARM_LowerArmLid->addButton(ui->rbLowerArmLid_NA, Repair_NA);

    // set ArmDriveInterface
    m_pgbRepairARM_ArmDriveInterface = new QButtonGroup( this);
    m_pgbRepairARM_ArmDriveInterface->addButton(ui->rbArmDriveInterface_OK, Repair_OK);
    m_pgbRepairARM_ArmDriveInterface->addButton(ui->rbArmDriveInterface_Repair, Repair_Repair);
    m_pgbRepairARM_ArmDriveInterface->addButton(ui->rbArmDriveInterface_NA, Repair_NA);

    // set ArmGripperInterfaceScara
    m_pgbRepairARM_ArmGripperInterfaceScara = new QButtonGroup( this);
    m_pgbRepairARM_ArmGripperInterfaceScara->addButton(ui->rbArmGripperInterfaceScara_OK, Repair_OK);
    m_pgbRepairARM_ArmGripperInterfaceScara->addButton(ui->rbArmGripperInterfaceScara_Repair, Repair_Repair);
    m_pgbRepairARM_ArmGripperInterfaceScara->addButton(ui->rbArmGripperInterfaceScara_NA, Repair_NA);

    // set ArmGripperInterfaceNT
    m_pgbRepairARM_ArmGripperInterfaceNT = new QButtonGroup( this);
    m_pgbRepairARM_ArmGripperInterfaceNT->addButton(ui->rbArmGripperInterfaceScara_OK, Repair_OK);
    m_pgbRepairARM_ArmGripperInterfaceNT->addButton(ui->rbArmGripperInterfaceScara_Repair, Repair_Repair);
    m_pgbRepairARM_ArmGripperInterfaceNT->addButton(ui->rbArmGripperInterfaceScara_NA, Repair_NA);

    // set BeltReel
    m_pgbRepairARM_BeltReel = new QButtonGroup( this);
    m_pgbRepairARM_BeltReel->addButton(ui->rbBeltReel_OK, Repair_OK);
    m_pgbRepairARM_BeltReel->addButton(ui->rbBeltReel_Repair, Repair_Repair);
    m_pgbRepairARM_BeltReel->addButton(ui->rbBeltReel_NA, Repair_NA);

    // set TorxScrew
    m_pgbRepairARM_TorxScrew = new QButtonGroup( this);
    m_pgbRepairARM_TorxScrew->addButton(ui->rbTorxScrew_OK, Repair_OK);
    m_pgbRepairARM_TorxScrew->addButton(ui->rbTorxScrew_Repair, Repair_Repair);
    m_pgbRepairARM_TorxScrew->addButton(ui->rbTorxScrew_NA, Repair_NA);

    // set Bearings
    m_pgbRepairARM_Bearings = new QButtonGroup( this);
    m_pgbRepairARM_Bearings->addButton(ui->rbBearings_OK, Repair_OK);
    m_pgbRepairARM_Bearings->addButton(ui->rbBearings_Repair, Repair_Repair);
    m_pgbRepairARM_Bearings->addButton(ui->rbBearings_NA, Repair_NA);

    // set DeliverTo
    m_pgbRepairARM_Deliver = new QButtonGroup( this);
    m_pgbRepairARM_Deliver->addButton(ui->cbArmTW, TW);
    m_pgbRepairARM_Deliver->addButton(ui->cbArmEU, EU);

}

void CAnalyseData::setRadioButtonsIDsInGB4RepairDM( void){
    // set DMLikaMotor
    m_pgbRepairDM_DMLikaMotor = new QButtonGroup( this);
    m_pgbRepairDM_DMLikaMotor->addButton(ui->rbDMLikaMotor_OK, Repair_OK);
    m_pgbRepairDM_DMLikaMotor->addButton(ui->rbDMLikaMotor_Repair, Repair_Repair);
    m_pgbRepairDM_DMLikaMotor->addButton(ui->rbDMLikaMotor_NA, Repair_NA);

    // set CableHood
    m_pgbRepairDM_CableHood = new QButtonGroup( this);
    m_pgbRepairDM_CableHood->addButton(ui->rbCableHood_OK, Repair_OK);
    m_pgbRepairDM_CableHood->addButton(ui->rbCableHood_Repair, Repair_Repair);
    m_pgbRepairDM_CableHood->addButton(ui->rbCableHood_NA, Repair_NA);

    // set DMHousing
    m_pgbRepairDM_DMHousing = new QButtonGroup( this);
    m_pgbRepairDM_DMHousing->addButton(ui->rbDMHousing_OK, Repair_OK);
    m_pgbRepairDM_DMHousing->addButton(ui->rbDMHousing_Repair, Repair_Repair);
    m_pgbRepairDM_DMHousing->addButton(ui->rbDMHousing_NA, Repair_NA);

    // set DMLid
    m_pgbRepairDM_DMLid = new QButtonGroup( this);
    m_pgbRepairDM_DMLid->addButton(ui->rbDMLid_OK, Repair_OK);
    m_pgbRepairDM_DMLid->addButton(ui->rbDMLid_Repair, Repair_Repair);
    m_pgbRepairDM_DMLid->addButton(ui->rbDMLid_NA, Repair_NA);

    // set SlipRing
    m_pgbRepairDM_SlipRing = new QButtonGroup( this);
    m_pgbRepairDM_SlipRing->addButton(ui->rbSlipRing_OK, Repair_OK);
    m_pgbRepairDM_SlipRing->addButton(ui->rbSlipRing_Repair, Repair_Repair);
    m_pgbRepairDM_SlipRing->addButton(ui->rbSlipRing_NA, Repair_NA);

    // set HollowShaft
    m_pgbRepairDM_HollowShaft = new QButtonGroup( this);
    m_pgbRepairDM_HollowShaft->addButton(ui->rbHollowShaft_OK, Repair_OK);
    m_pgbRepairDM_HollowShaft->addButton(ui->rbHollowShaft_Repair, Repair_Repair);
    m_pgbRepairDM_HollowShaft->addButton(ui->rbHollowShaft_NA, Repair_NA);

    // set DeliverTo
    m_pgbRepairDM_Deliver = new QButtonGroup( this);
    m_pgbRepairDM_Deliver->addButton(ui->cbDMTW, TW);
    m_pgbRepairDM_Deliver->addButton(ui->cbDMEU, EU);
}

void CAnalyseData::setRadioButtonsIDsInGB4RepairZT( void){
    // set ZStroke35
    m_pgbRepairZT_ZStroke35 = new QButtonGroup( this);
    m_pgbRepairZT_ZStroke35->addButton(ui->rbZStroke35_OK, Repair_OK);
    m_pgbRepairZT_ZStroke35->addButton(ui->rbZStroke35_Repair, Repair_Repair);
    m_pgbRepairZT_ZStroke35->addButton(ui->rbZStroke35_NA, Repair_NA);

    // set ZStroke50
    m_pgbRepairZT_ZStroke50 = new QButtonGroup( this);
    m_pgbRepairZT_ZStroke50->addButton(ui->rbZStroke50_OK, Repair_OK);
    m_pgbRepairZT_ZStroke50->addButton(ui->rbZStroke50_Repair, Repair_Repair);
    m_pgbRepairZT_ZStroke50->addButton(ui->rbZStroke50_NA, Repair_NA);

    // set ZMHousingScara
    m_pgbRepairZT_ZMHousingScara = new QButtonGroup( this);
    m_pgbRepairZT_ZMHousingScara->addButton(ui->rbZMHousingScara_OK, Repair_OK);
    m_pgbRepairZT_ZMHousingScara->addButton(ui->rbZMHousingScara_Repair, Repair_Repair);
    m_pgbRepairZT_ZMHousingScara->addButton(ui->rbZMHousingScara_NA, Repair_NA);

    // set ZMHousingNT
    m_pgbRepairZT_ZMHousingNT = new QButtonGroup( this);
    m_pgbRepairZT_ZMHousingNT->addButton(ui->rbZMHousingNT_OK, Repair_OK);
    m_pgbRepairZT_ZMHousingNT->addButton(ui->rbZMHousingNT_Repair, Repair_Repair);
    m_pgbRepairZT_ZMHousingNT->addButton(ui->rbZMHousingNT_NA, Repair_NA);

    // set GuidingShaftsScara
    m_pgbRepairZT_GuidingShaftsScara = new QButtonGroup( this);
    m_pgbRepairZT_GuidingShaftsScara->addButton(ui->rbGuidingShaftsScara_OK, Repair_OK);
    m_pgbRepairZT_GuidingShaftsScara->addButton(ui->rbGuidingShaftsScara_Repair, Repair_Repair);
    m_pgbRepairZT_GuidingShaftsScara->addButton(ui->rbGuidingShaftsScara_NA, Repair_NA);

    // set GuidingShaftsNT
    m_pgbRepairZT_GuidingShaftsNT = new QButtonGroup( this);
    m_pgbRepairZT_GuidingShaftsNT->addButton(ui->rbGuidingShaftsNT_OK, Repair_OK);
    m_pgbRepairZT_GuidingShaftsNT->addButton(ui->rbGuidingShaftsNT_Repair, Repair_Repair);
    m_pgbRepairZT_GuidingShaftsNT->addButton(ui->rbGuidingShaftsNT_NA, Repair_NA);

    // set SmallGuidingShafts
    m_pgbRepairZT_SmallGuidingShafts = new QButtonGroup( this);
    m_pgbRepairZT_SmallGuidingShafts->addButton(ui->rbSmallGuidingShafts_OK, Repair_OK);
    m_pgbRepairZT_SmallGuidingShafts->addButton(ui->rbSmallGuidingShafts_Repair, Repair_Repair);
    m_pgbRepairZT_SmallGuidingShafts->addButton(ui->rbSmallGuidingShafts_NA, Repair_NA);

    // set ClampingFlange
    m_pgbRepairZT_ClampingFlange = new QButtonGroup( this);
    m_pgbRepairZT_ClampingFlange->addButton(ui->rbClampingFlange_OK, Repair_OK);
    m_pgbRepairZT_ClampingFlange->addButton(ui->rbClampingFlange_Repair, Repair_Repair);
    m_pgbRepairZT_ClampingFlange->addButton(ui->rbClampingFlange_NA, Repair_NA);

    // set AdapterCable
    m_pgbRepairZT_AdapterCable = new QButtonGroup( this);
    m_pgbRepairZT_AdapterCable->addButton(ui->rbAdapterCable_OK, Repair_OK);
    m_pgbRepairZT_AdapterCable->addButton(ui->rbAdapterCable_Repair, Repair_Repair);
    m_pgbRepairZT_AdapterCable->addButton(ui->rbAdapterCable_NA, Repair_NA);

    // set DeliverTo
    m_pgbRepairZT_Deliver = new QButtonGroup( this);
    m_pgbRepairZT_Deliver->addButton(ui->cbZTTW, TW);
    m_pgbRepairZT_Deliver->addButton(ui->cbZTEU, EU);
}

void CAnalyseData::writeAmount( QAxObject* workbook, sREPAIRITEM item )
{
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);
    int row = 0, column = 0;
    CAnalyseData::getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    // if checkbox "repair" is checked -> write "1"
    if(item.Value.toInt() == 1){
        QString value = "'1";
        cell->setProperty("Value", value);
    }
}

void CAnalyseData::writeInformation( QAxObject* workbook, sREPAIRITEM item )
{
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);
    int row = 0, column = 0;
    CAnalyseData::getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    // write information
    QString value = item.Value.toString();
    cell->setProperty("Value", value);
}

void CAnalyseData::extendInformation( QAxObject* workbook, sREPAIRITEM item )
{
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);
    int row = 0, column = 0;
    CAnalyseData::getRowColumn(item.Cell, &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    // write information

    QString value = cell->property("Value").toString() + " " + item.Value.toString();
    cell->setProperty("Value", value);
}

void CAnalyseData::buildRepairTable( void )
{
    // Value                                                                                // Cell     // Function
    // ----- General Information --------------------------------------------------------------------------
    vecRepairItems.append({"'"+ui->leNew12NC->text(),                                       "C4",       writeInformation});
    vecRepairItems.append({"'"+ui->leRobotSN->text(),                                       "C5",       writeInformation});
    vecRepairItems.append({"'"+ui->leARMSN->text(),                                         "D7",       writeInformation});
    vecRepairItems.append({ui->leArmFirstDelivery->text()+ " ; "+ui->leArmLastRepair->text(),"B7",       extendInformation});
    vecRepairItems.append({((ui->leDMSN->text()).split("-"))[3],                            "D8",       writeInformation});
    vecRepairItems.append({ui->leDMFirstDelivery->text()+" ; "+ui->leDMLastRepair->text(),   "B8",       extendInformation});
    vecRepairItems.append({((ui->leZTSN2->text()).split("-"))[3],                           "D9",       writeInformation});
    vecRepairItems.append({ui->leZTFirstDelivery->text()+" ; "+ui->leZTLastRepair->text(),   "B9",       extendInformation});
    vecRepairItems.append({ui->leRepairNo->text()+ " ; " +ui->leLastRepairDate->text(),      "F5",       extendInformation});
    vecRepairItems.append({ui->leFirstDeliveryDate->text(),                                 "F6",       extendInformation});

    // ------ Arm parts ------------------------------------------------------------------------------------
    vecRepairItems.append({m_pgbRepairARM_ArmBelts->checkedId(),                            "E13",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_UpperArmHousingUpgrade->checkedId(),              "E35",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_UpperArmHousing->checkedId(),                     "E36",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_UpperArmLid->checkedId(),                         "E37",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_LowerArmHousingUpgrade->checkedId(),              "E38",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_LowerArmHousing->checkedId(),                     "E39",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_LowerArmLid->checkedId(),                         "E40",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_ArmDriveInterface->checkedId(),                   "E41",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_ArmGripperInterfaceScara->checkedId(),            "E42",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_ArmGripperInterfaceNT->checkedId(),               "E43",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_BeltReel->checkedId(),                            "E44",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_TorxScrew->checkedId(),                           "E45",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_Bearings->checkedId(),                            "E46",      writeAmount});
    vecRepairItems.append({m_pgbRepairARM_Deliver->checkedId(),                             "E47",      writeAmount});

////    // ------ DM parts ------------------------------------------------------------------------------------
    vecRepairItems.append({m_pgbRepairDM_DMLikaMotor->checkedId(),                          "E15",      writeAmount});
    vecRepairItems.append({m_pgbRepairDM_CableHood->checkedId(),                            "E49",      writeAmount});
    vecRepairItems.append({m_pgbRepairDM_DMHousing->checkedId(),                            "E50",      writeAmount});
    vecRepairItems.append({m_pgbRepairDM_DMLid->checkedId(),                                "E51",      writeAmount});
    vecRepairItems.append({m_pgbRepairDM_SlipRing->checkedId(),                             "E52",      writeAmount});
    vecRepairItems.append({m_pgbRepairDM_HollowShaft->checkedId(),                          "E53",      writeAmount});
    vecRepairItems.append({m_pgbRepairDM_Deliver->checkedId(),                              "E54",      writeAmount});

////    // ------ Z-module parts ------------------------------------------------------------------------------
    vecRepairItems.append({m_pgbRepairZT_ZStroke35->checkedId(),                            "E17",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_ZStroke50->checkedId(),                            "E18",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_ZMHousingScara->checkedId(),                       "E56",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_ZMHousingNT->checkedId(),                          "E57",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_GuidingShaftsScara->checkedId(),                   "E58",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_GuidingShaftsNT->checkedId(),                      "E59",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_SmallGuidingShafts->checkedId(),                   "E60",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_ClampingFlange->checkedId(),                       "E61",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_AdapterCable->checkedId(),                         "E62",      writeAmount});
    vecRepairItems.append({m_pgbRepairZT_Deliver->checkedId(),                              "E63",      writeAmount});

    // ------ Upgrade part --------------------------------------------------------------------------------

    // ------ Labour & Packaging --------------------------------------------------------------------------
}

void CAnalyseData::getDataFromRepair( void ){
    getRepairFromARM();
    getRepairFromDM();
    getRepairFromZT();
}
void CAnalyseData::getRepairFromARM( void )
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMREPAIR);

    QString strInputName = "ArmBelts";
    switch( m_pgbRepairARM_ArmBelts->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "UpperArmHousingUpgrade";
    switch( m_pgbRepairARM_UpperArmHousingUpgrade->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "UpperArmHousing";
    switch( m_pgbRepairARM_UpperArmHousing->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "UpperArmLid";
    switch( m_pgbRepairARM_UpperArmLid->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "LowerArmHousingUpgrade";
    switch( m_pgbRepairARM_LowerArmHousingUpgrade->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "LowerArmHousing";
    switch( m_pgbRepairARM_LowerArmHousing->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "LowerArmLid";
    switch( m_pgbRepairARM_LowerArmLid->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ArmDriveInterface";
    switch( m_pgbRepairARM_ArmDriveInterface->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ArmGripperInterfaceScara";
    switch( m_pgbRepairARM_ArmGripperInterfaceScara->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ArmGripperInterfaceNT";
    switch( m_pgbRepairARM_ArmGripperInterfaceNT->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "BeltReel";
    switch( m_pgbRepairARM_BeltReel->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "TorxScrew";
    switch( m_pgbRepairARM_TorxScrew->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "Bearings";
    switch( m_pgbRepairARM_Bearings->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "Deliver";
    switch( m_pgbRepairARM_Deliver->checkedId())
    {
        case TW:
        settings.setValue(strInputName, TW);
        break;
        case EU:
        settings.setValue(strInputName, EU);
        break;
    }


}
void CAnalyseData::getRepairFromDM( void )
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMREPAIR);

    QString strInputName = "DMLikaMotor";
    switch( m_pgbRepairDM_DMLikaMotor->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "CableHood";
    switch( m_pgbRepairDM_CableHood->checkedId())
    {
        case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
        case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
        case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "DMLid";
    switch( m_pgbRepairDM_DMLid->checkedId())
    {
        case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
        case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
        case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "SlipRing";
    switch( m_pgbRepairDM_SlipRing->checkedId())
    {
        case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
        case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
        case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "HollowShaft";
    switch( m_pgbRepairDM_HollowShaft->checkedId())
    {
        case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
        case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
        case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "Deliver";
    switch( m_pgbRepairDM_Deliver->checkedId())
    {
        case TW:
        settings.setValue(strInputName, TW);
        break;
        case EU:
        settings.setValue(strInputName, EU);
        break;
    }
}
void CAnalyseData::getRepairFromZT( void )
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTREPAIR);

    QString strInputName = "ZStroke35";
    switch( m_pgbRepairZT_ZStroke35->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ZStroke50";
    switch( m_pgbRepairZT_ZStroke50->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ZMHousingScara";
    switch( m_pgbRepairZT_ZMHousingScara->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ZMHousingNT";
    switch( m_pgbRepairZT_ZMHousingNT->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "GuidingShaftsScara";
    switch( m_pgbRepairZT_GuidingShaftsScara->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "GuidingShaftsNT";
    switch( m_pgbRepairZT_GuidingShaftsNT->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "SmallGuidingShafts";
    switch( m_pgbRepairZT_SmallGuidingShafts->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "ClampingFlange";
    switch( m_pgbRepairZT_ClampingFlange->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "AdapterCable";
    switch( m_pgbRepairZT_AdapterCable->checkedId())
    {
      case Repair_OK:
        settings.setValue(strInputName, Repair_OK);
        break;
      case Repair_Repair:
        settings.setValue(strInputName, Repair_Repair);
        break;
      case Repair_NA:
        settings.setValue(strInputName, Repair_NA);
        break;
    }
    strInputName = "Deliver";
    switch( m_pgbRepairZT_Deliver->checkedId())
    {
        case TW:
        settings.setValue(strInputName, TW);
        break;
        case EU:
        settings.setValue(strInputName, EU);
        break;
    }
}

void CAnalyseData::getDataFromIni4Repair( void ){
    getDataFromIni4ARMRepair();
    getDataFromIni4DMRepair();
    getDataFromIni4ZTRepair();
}
void CAnalyseData::getDataFromIni4ARMRepair( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ARMREPAIR);

    QString strInputName = "ArmBelts";
    int iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbArmBeltsAvrPrice_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbArmBeltsAvrPrice_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbArmBeltsAvrPrice_NA->setChecked( true);
        break;
    }
    strInputName = "UpperArmHousingUpgrade";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbUpperArmHousingUpgrade_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbUpperArmHousingUpgrade_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbUpperArmHousingUpgrade_NA->setChecked( true);
        break;
    }
    strInputName = "UpperArmHousing";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbUpperArmHousing_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbUpperArmHousing_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbUpperArmHousing_NA->setChecked( true);
        break;
    }
    strInputName = "UpperArmLid";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbUpperArmLid_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbUpperArmLid_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbUpperArmLid_NA->setChecked( true);
        break;
    }
    strInputName = "LowerArmHousingUpgrade";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbLowerArmHousingUpgrade_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbLowerArmHousingUpgrade_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbLowerArmHousingUpgrade_NA->setChecked( true);
        break;
    }
    strInputName = "LowerArmHousing";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbLowerArmHousing_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbLowerArmHousing_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbLowerArmHousing_NA->setChecked( true);
        break;
    }
    strInputName = "LowerArmLid";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbLowerArmLid_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbLowerArmLid_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbLowerArmLid_NA->setChecked( true);
        break;
    }
    strInputName = "ArmDriveInterfac";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbArmDriveInterface_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbArmDriveInterface_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbArmDriveInterface_NA->setChecked( true);
        break;
    }
    strInputName = "ArmGripperInterfaceScara";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbArmGripperInterfaceScara_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbArmGripperInterfaceScara_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbArmGripperInterfaceScara_NA->setChecked( true);
        break;
    }
    strInputName = "ArmGripperInterfaceNT";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbArmGripperInterfaceNT_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbArmGripperInterfaceNT_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbArmGripperInterfaceNT_NA->setChecked( true);
        break;
    }
    strInputName = "BeltReel";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbBeltReel_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbBeltReel_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbBeltReel_NA->setChecked( true);
        break;
    }
    strInputName = "TorxScrew";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbTorxScrew_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbTorxScrew_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbTorxScrew_NA->setChecked( true);
        break;
    }
    strInputName = "Bearings";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbBearings_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbBearings_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbBearings_NA->setChecked( true);
        break;
    }
    strInputName = "Deliver";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case TW:
        ui->cbArmTW->setChecked(true);
        break;
      case EU:
        ui->cbArmEU->setChecked( true);
        break;
    }
}
void CAnalyseData::getDataFromIni4DMRepair( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_DMREPAIR);

    QString strInputName = "DMLikaMotor";
    int iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbDMLikaMotor_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbDMLikaMotor_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbDMLikaMotor_NA->setChecked( true);
        break;
    }
    strInputName = "CableHood";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbCableHood_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbCableHood_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbCableHood_NA->setChecked( true);
        break;
    }
    strInputName = "DMHousing";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbDMHousing_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbDMHousing_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbDMHousing_NA->setChecked( true);
        break;
    }
    strInputName = "DMLid";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbDMLid_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbDMLid_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbDMLid_NA->setChecked( true);
        break;
    }
    strInputName = "SlipRing";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbSlipRing_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbSlipRing_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbSlipRing_NA->setChecked( true);
        break;
    }
    strInputName = "HollowShaft";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbHollowShaft_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbHollowShaft_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbHollowShaft_NA->setChecked( true);
        break;
    }
    strInputName = "Deliver";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case TW:
        ui->cbDMTW->setChecked( true);
        break;
      case EU:
        ui->cbDMEU->setChecked( true);
        break;
    }
}
void CAnalyseData::getDataFromIni4ZTRepair( void)
{
    QSettings settings(m_filePath,QSettings::IniFormat);
    settings.beginGroup(BEGIN_ZTREPAIR);

    QString strInputName = "ZStroke35";
    int iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbZStroke35_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbZStroke35_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbZStroke35_NA->setChecked( true);
        break;
    }
    strInputName = "ZStroke50";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbZStroke50_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbZStroke50_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbZStroke50_NA->setChecked( true);
        break;
    }
    strInputName = "ZMHousingScara";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbZMHousingScara_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbZMHousingScara_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbZMHousingScara_NA->setChecked( true);
        break;
    }
    strInputName = "ZMHousingNT";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbZMHousingNT_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbZMHousingNT_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbZMHousingNT_NA->setChecked( true);
        break;
    }
    strInputName = "GuidingShaftsScara";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbGuidingShaftsScara_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbGuidingShaftsScara_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbGuidingShaftsScara_NA->setChecked( true);
        break;
    }
    strInputName = "GuidingShaftsNT";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbGuidingShaftsNT_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbGuidingShaftsNT_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbGuidingShaftsNT_NA->setChecked( true);
        break;
    }
    strInputName = "SmallGuidingShafts";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case Repair_OK:
        ui->rbSmallGuidingShafts_OK->setChecked( true);
        break;
      case Repair_Repair:
        ui->rbSmallGuidingShafts_Repair->setChecked( true);
        break;
      case Repair_NA:
        ui->rbSmallGuidingShafts_NA->setChecked( true);
        break;
    }
    strInputName = "Deliver";
    iType = settings.value(strInputName).toInt();
    switch(iType)
    {
      case TW:
        ui->cbZTTW->setChecked(true);
        break;
      case EU:
        ui->cbZTEU->setChecked( true);
        break;
    }

}

void CAnalyseData::createRepairMatrix( void )
{

//    QString m_robotNumber = ui->lbRobotType->text() + ui->leRobotTypeSN->text();
//    m_filePathExcelRepair = "D:\\Data\\twintern\\Jana\\Work\\Emily\\Files\\RepairSheet_w.xlsx";  //
//    m_filePathExcelRepair.append(m_robotType + ui->leRobotTypeSN->text() +"_w.xlsx"); // strExcelFilePath + "\\" + strExcelRepairTemp + "_" + m_robotNumber +"_w.xlsx";

//    m_filePathExcelRepairTmp = "D:\\Data\\twintern\\Jana\\Work\\Emily\\Files\\Repair_matrix _MK5.xlsx";

    ui->lbStatus->setText("Creating Repair Matrix");
    progressSave(3);
    closeExcel();
    progressSave(6);
    progressSave(10);
    buildRepairTable();
    if(QFile::exists(m_filePathExcelRepair))
    {
        QFile::remove(m_filePathExcelRepair);
    }

    m_objExcel = new QAxObject("Excel.Application");
    if( m_objExcel==nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Excel is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }

    m_objExcel->setProperty("Visible",false);
    QAxObject* workbooks = m_objExcel->querySubObject("WorkBooks");// get the workbook.
    if( workbooks==nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Office is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }
    workbooks->setProperty("Visible",false);
    workbooks->dynamicCall("Open (const QString&)", m_filePathExcelRepairTmp);   // Open the file；
    m_objWorkbook = m_objExcel->querySubObject("ActiveWorkBook"); // Get the active workbook.
    progressSave(20);
    int item_count = 1;// for progress
    foreach(sREPAIRITEM item, vecRepairItems)
    {
        if(item.p_func)
        {
            (this->*(item.p_func))(m_objWorkbook, item);
        }
        progressSave(20 + (item_count++)*80/vecRepairItems.size());
    }
    // add to disable checking compatibility
    m_objWorkbook->setProperty("DisplayAlerts", false);
    m_objWorkbook->setProperty("CheckCompatibility", false);
    m_objWorkbook->setProperty("DoNotPromptForConvert", true);

    m_objWorkbook->dynamicCall("SaveAs(const QString&)",
                                     QDir::toNativeSeparators(m_filePathExcelRepair));

    progressSave(98);

    closeExcel();
    progressSave(100);
    qDebug() << "Repair Matrix created.";
    ui->lbStatus->setText("Repair Matrix was created");

}

void CAnalyseData::on_pbExport_clicked( void )
{
    ui->lbStatus->setText("Starting export");

    // track progress
    int counterAll=0;
    int counterItem=0;
    if(ui->cbAnalyseSheet->isChecked())
        counterAll++;
    if(ui->cbRepairMatrix->isChecked())
        counterAll++;
    if(ui->cbPrintLabel->isChecked())
        counterAll++;
    if(ui->cbMOMSheet->isChecked())
        counterAll++;

    ui->lbProcess->setText(QString::number(counterItem) +"/"+QString::number(counterAll));
    qDebug() << counterAll;

    if(counterAll > 0){
        progressSave(3);
        on_pbSaveIni_clicked();
        progressSave(6);
    }else{
        progressSave(3);
        on_pbSaveIni_clicked();
        progressSave(0);
        ui->lbStatus->setText("WARNING: No export selected! Ini file was still created.");
        return;
    }

    if(ui->cbAnalyseSheet->isChecked()){
        // export Analyse Excel
        createAnalyseSheet();
        ui->lbProcess->setText(QString::number(++counterItem) +"/"+QString::number(counterAll));
    }
    if(ui->cbRepairMatrix->isChecked()){
        // create Repair Matrix
        createRepairMatrix();
        ui->lbProcess->setText(QString::number(++counterItem) +"/"+QString::number(counterAll));
    }
    if(ui->cbPrintLabel->isChecked()){
        createPrintLabel();
        ui->lbProcess->setText(QString::number(++counterItem) +"/"+QString::number(counterAll));
    }
    if(ui->cbMOMSheet->isChecked()){
        createMOMSheet();
        ui->lbProcess->setText(QString::number(++counterItem) +"/"+ QString::number(counterAll));
    }
    ui->lbStatus->setText("Export finished");

}


// ************** Create MOM Sheet -----------------------------

void CAnalyseData::createMOMSheet( void ){

//    on_pbSaveIni_clicked();
    ui->lbStatus->setText("Creating MOM Sheet");
    closeExcel();
    progressSave(3);
    buildProtocolTable();
    if(QFile::exists(m_filePathExcelMOM))
    {
        QFile::remove(m_filePathExcelMOM);
    }

    m_objExcel = new QAxObject("Excel.Application");
    if( m_objExcel==nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Excel is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }

    m_objExcel->setProperty("Visible",false);
    QAxObject* workbooks = m_objExcel->querySubObject("WorkBooks");// get the workbook.
    if( workbooks==nullptr)
    {
        progressSave(0);
        QMessageBox::critical(NULL, "Error", "Office is not installed", QMessageBox::Yes, QMessageBox::Yes);
        return;
    }
    workbooks->setProperty("Visible",false);    //new code*********************
    workbooks->dynamicCall("Open (const QString&)", m_filePathExcelMOMTmp);   // Open the file；
    m_objWorkbook = m_objExcel->querySubObject("ActiveWorkBook"); // Get the active workbook.
    progressSave(20);
    QAxObject *worksheet = m_objWorkbook->querySubObject("WorkSheets(int)", 1);

    int row = 0, column = 0;

    // Write Robot Number
    CAnalyseData::getRowColumn("A2", &row, &column);
    QAxObject* cell = worksheet->querySubObject("Cells(int,int)", row, column);
    QString value = m_robotType + m_robotSN;
    cell->setProperty("Value", value);

    // Write MN (repair number)
    column++;
    cell = worksheet->querySubObject("Cells(int,int)", row, column);
    value = ui->leRepairNRARM->text();
    cell->setProperty("Value", value);

    // Write WK
    QStringList parts = m_fileName.split("_");
    if(parts.size() > 2) {
      value = parts[2];
    }
    column++;
    cell = worksheet->querySubObject("Cells(int,int)", row, column);
    cell->setProperty("Value", value);

    // Write Result Analysis
    column++;
    cell = worksheet->querySubObject("Cells(int,int)", row, column);
    strResAnalysis = cell->property("Value").toString();
    startPos = 0;

    progressSave(30);
    writeGeneralInfoMOM();
    progressSave(40);
    writeArmInfoMOM();
    progressSave(60);
    writeDMInfoMOM();
    progressSave(80);
    writeZInfoMOM();
    progressSave(95);

    cell->setProperty("Value", strResAnalysis);
    qDebug() << strResAnalysis;

    // make headers bold:
    int idx = strResAnalysis.indexOf("SA:");
    QAxObject* chars = cell->querySubObject("Characters(int, int)", idx, 3);
    QAxObject *font = chars->querySubObject("Font");
    font->setProperty("Bold", true);
    idx = strResAnalysis.indexOf("DM:");
    chars = cell->querySubObject("Characters(int, int)", idx, 3);
    font = chars->querySubObject("Font");
    font->setProperty("Bold", true);
    idx = strResAnalysis.indexOf("Z:");
    chars = cell->querySubObject("Characters(int, int)", idx, 3);
    font = chars->querySubObject("Font");
    font->setProperty("Bold", true);

    m_objWorkbook->setProperty("DisplayAlerts", false);
    m_objWorkbook->setProperty("CheckCompatibility", false);
    m_objWorkbook->setProperty("DoNotPromptForConvert", true);

    m_objWorkbook->dynamicCall("SaveAs(const QString&)",
                                     QDir::toNativeSeparators(m_filePathExcelMOM));

    progressSave(98);

    closeExcel();
    progressSave(100);
    qDebug() << "MOM Sheet created.";
    ui->lbStatus->setText("MOM Sheet was created");

}

void CAnalyseData::writeGeneralInfoMOM(){

    getNextLineIdx();
    // Version
    switch(m_pgbHDMotorType->checkedId())
    {
      case V0:
        insertAndReturnLastIdx(": V0");
        break;
      case V1:
        insertAndReturnLastIdx(": V1");
        break;
      case DFV1:
        insertAndReturnLastIdx(": DFV1");
        break;
    }
    getNextLineIdx();
    // First Delivery
    insertAndReturnLastIdx(ui->leFirstDeliveryDate->text());
    getNextLineIdx();
    //Repair No
    insertAndReturnLastIdx(ui->leRepairNo->text());
    getNextLineIdx();
    // Last repair date
    if(!(ui->leLastRepairDate->text()).isEmpty()){
        removeAndReturnLastIdx(2);
        insertAndReturnLastIdx(ui->leLastRepairDate->text());
    }
    getNextLineIdx(); // "Delivered in white box\n"
    getNextLineIdx(); // "packed regarding spec: No, ARM down holder placed incorrectly\n"
    getNextLineIdx(); // "Delivered in white box\n"
    getNextLineIdx(); // "Unit starts on testing: OK\n"
    getNextLineIdx(); // "/n"


}
void CAnalyseData::writeArmInfoMOM(){
    getNextLineIdx(); // "SA:/n"
    // "- Geometry: \n"
    switch(m_pgbGeoChkARM->checkedId())
    {
      case Test_OK:
        insertAndReturnLastIdx("OK");
        break;
      case Test_NG:
        insertAndReturnLastIdx("NOK");
        break;
      case Test_NA:
        insertAndReturnLastIdx("N/A");
        break;
    }
    getNextLineIdx();
    // Rz
    double dRz = ui->leGeoRz->text().toDouble();
    if(qAbs(dRz)> Rz_UP){
        startPos-=21;
        insertAndReturnLastIdx(ui->leGeoRz->text()+" mRad ");
        getNextLineIdx();
    }else{
        removeAndReturnLastIdx(26);
//        getNextLineIdx();
    }
    getNextLineIdx();
    // Rx
    double dRx = ui->leGeoRx->text().toDouble();
    if(qAbs(dRx)> Rx_UP){
        startPos-=21;
        insertAndReturnLastIdx(ui->leGeoRx->text()+" mRad ");
        getNextLineIdx();
    }else{
        removeAndReturnLastIdx(26);
//        getNextLineIdx();
    }
    getNextLineIdx();
    // Ry
    double dRy = ui->leGeoRy->text().toDouble();
    if(qAbs(dRy)> Ry_UP){
        startPos-=21;
        insertAndReturnLastIdx(ui->leGeoRy->text()+" mRad ");
        getNextLineIdx();
    }else{
        removeAndReturnLastIdx(26);
//        getNextLineIdx();
    }
    getNextLineIdx();

    // delta H4
    double dH4 = ui->leGeoRy->text().toDouble();
    if(dH4 < DeltaHeight_DN && dH4 > DeltaHeight_UP){
        startPos-=29;
        insertAndReturnLastIdx(ui->leGeoDelHeight->text()+" mm ");
        getNextLineIdx();
    }else{
        removeAndReturnLastIdx(35);
//        getNextLineIdx();
    }

    getNextLineIdx();
    // Position TH

    double posR = ui->leRepPosPAR->text().toDouble();
    double posTH = ui->leRepPosPATH->text().toDouble();
    qDebug() << QString::number(posR);
    qDebug() << QString::number(posTH);
    if(qAbs(posR) > REPPOSPA_UP && qAbs(posTH) > REPPOSPA_UP){
        startPos-=22;
        insertAndReturnLastIdx(QString::number(posTH) +" µm (spec.= +/- 600 µm), Pos_R="+QString::number(posR)+" ");
        getNextLineIdx();
    }else if(qAbs(posTH) > REPPOSPA_UP && qAbs(posR) <= REPPOSPA_UP){
        startPos-=22;
        insertAndReturnLastIdx(QString::number(posTH)+" ");
        getNextLineIdx();
    }else if(qAbs(posTH) <= REPPOSPA_UP && qAbs(posR) > REPPOSPA_UP){
        startPos-=22;
        removeAndReturnLastIdx(4);
        insertAndReturnLastIdx("R="+QString::number(posR)+" ");
        getNextLineIdx();
    }else{
        removeAndReturnLastIdx(30);
        insertAndReturnLastIdx("OK");
    }
    getNextLineIdx();
    // "- Electrics ARM: \n"
    switch(m_pgbEleChkARM->checkedId())
    {
      case Test_OK:
        insertAndReturnLastIdx(" OK");
        break;
      case Test_NG:
        insertAndReturnLastIdx(" NOK");
        break;
      case Test_NA:
        insertAndReturnLastIdx(" N/A");
        break;
    }
    getNextLineIdx();
    getNextLineIdx(); //"Visual:\n"
    getNextLineIdx(); // "\n"
}
void CAnalyseData::writeDMInfoMOM(){
    getNextLineIdx(); // "DM:\n"
    // Geometry 180
    double geo180 = ui->le180DegVal_2->text().toDouble();
    double geo270 = ui->le270DegVal_2->text().toDouble();
    if(qAbs(geo180) > MicroHite_UP){
        startPos-=44;
        insertAndReturnLastIdx(ui->le180DegVal->text()+" ");
        getNextLineIdx();
    }else{
        startPos-=38;
        removeAndReturnLastIdx(16);
        getNextLineIdx();
    }
    // Geometry 270
    if(qAbs(geo270) > MicroHite_UP){
        startPos-=26;
        insertAndReturnLastIdx(ui->le270DegVal->text()+" ");
        getNextLineIdx();
    }else{
        startPos-=22;
        removeAndReturnLastIdx(16);
        getNextLineIdx();
    }
    if(qAbs(geo180) <= MicroHite_UP && qAbs(geo270) <= MicroHite_UP){
        removeAndReturnLastIdx(21);
        insertAndReturnLastIdx("OK");
    }
    getNextLineIdx();
    // 0-Positioning TH
    double posTH = ui->leZeroingPosTH->text().toDouble();
    double posR = ui->leZeroingPosR->text().toDouble();
    qDebug() << "0TH: "<< posTH;
    qDebug() << "0R: "<< posR;
    if(qAbs(posTH)>0.1 && qAbs(posR)>0.1){
        startPos-=19;
        insertAndReturnLastIdx("Pos_TH="+ui->leZeroingPosTH->text()+" mm, Pos_R="+ui->leZeroingPosR->text()+" mm ");
        getNextLineIdx();
    }else if(qAbs(posTH)<= 0.1 && qAbs(posR)>0.1){
        startPos-=19;
        insertAndReturnLastIdx("Pos_R="+ui->leZeroingPosR->text()+" mm ");
        getNextLineIdx();
    }else if(qAbs(posTH)>0.1 && qAbs(posR)<=0.1){
        startPos-=19;
        insertAndReturnLastIdx("Pos_TH="+ui->leZeroingPosTH->text()+" mm ");
        getNextLineIdx();
    }else {
        removeAndReturnLastIdx(19);
        insertAndReturnLastIdx("OK");
    }

    getNextLineIdx();
    // Electrics DM:
    switch(m_pgbEleChkDM->checkedId())
    {
      case Test_OK:
        insertAndReturnLastIdx(" OK");
        break;
      case Test_NG:
        insertAndReturnLastIdx(" NOK");
        break;
      case Test_NA:
        insertAndReturnLastIdx(" N/A");
        break;
    }
    getNextLineIdx(); // "Visual:\n"
    getNextLineIdx(); // "- Motor / Encoder: \n"
    int motor1=m_pgbTHMotorChkDM->checkedId();
    int motor2=m_pgbRMotorChkDM->checkedId();
    if(motor1==Test_OK && motor2 ==Test_OK){
        insertAndReturnLastIdx(" OK");
    }else if(motor1==Test_NG || motor2 ==Test_NG){
        insertAndReturnLastIdx(" NOK");
    }else if(motor1==Test_NA || motor2 ==Test_NA){
        insertAndReturnLastIdx(" N/A");
    }
    getNextLineIdx();
    getNextLineIdx(); // "\n"
}
void CAnalyseData::writeZInfoMOM(){
    qDebug() << strResAnalysis;
    getNextLineIdx(); // "Z: \n"
    getNextLineIdx(); // "Positioning: \n"
    getNextLineIdx(); // "Electrics: \n"
    getNextLineIdx(); // "Visual: \n"
    getNextLineIdx(); // "Spindle/Motor:"
    // Electrics ZT:
    startPos=strResAnalysis.length();
    switch(m_pgbZMotorZT->checkedId())
    {
      case Test_OK:
        insertAndReturnLastIdx(" OK");
        break;
      case Test_NG:
        insertAndReturnLastIdx(" NOK");
        break;
      case Test_NA:
        insertAndReturnLastIdx(" N/A");
        break;
    }
}

void CAnalyseData::getNextLineIdx (){
    startPos++;
    startPos = strResAnalysis.indexOf("\n", startPos);
}
void CAnalyseData::insertAndReturnLastIdx(QString insertText){
    strResAnalysis.insert(startPos, insertText);
    startPos += insertText.length();
}
void CAnalyseData::removeAndReturnLastIdx(int numberRemove){
    strResAnalysis.remove(startPos-numberRemove, numberRemove);
    startPos-=numberRemove;
}
