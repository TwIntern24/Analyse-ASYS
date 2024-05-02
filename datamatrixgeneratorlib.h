#ifndef DATAMATRIXGENERATORLIB_H
#define DATAMATRIXGENERATORLIB_H

#include "datamatrixgeneratorlib_global.h"

#include <QDir>
#include <QMessageBox>
#include <QImage>
#include <algorithm>
#include <cctype>
#include <cstring>
#include <fstream>
#include <iostream>
#include <string>

#include "DataMatrix_ZXing/DMWriter.h"
#include "DataMatrix_ZXing/BitMatrix.h"

class DATAMATRIXGENERATORLIBSHARED_EXPORT DataMatrixGeneratorLib
{

public:
    DataMatrixGeneratorLib();

    int iMinWidth; int iMinHeight;
    int iMaxWidth; int iMaxHeight;
    int iMargin;

    ZXing::DataMatrix::SymbolShape _shapeHint;
    ZXing::DataMatrix::Writer DataMatrixWriter = ZXing::DataMatrix::Writer();

    void GenerateDataMatrix(QString strText, QString strSavePath, QString strFileName, int iFileType, int iSize);
};

#endif // DATAMATRIXGENERATORLIB_H
