/********************************************************************************
** Form generated from reading UI file 'AnalysisInputDialog.ui'
**
** Created by: Qt User Interface Compiler version 4.8.7
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_ANALYSISINPUTDIALOG_H
#define UI_ANALYSISINPUTDIALOG_H

#include <QtCore/QVariant>
#include <QtGui/QAction>
#include <QtGui/QApplication>
#include <QtGui/QButtonGroup>
#include <QtGui/QComboBox>
#include <QtGui/QDialog>
#include <QtGui/QDialogButtonBox>
#include <QtGui/QDoubleSpinBox>
#include <QtGui/QFormLayout>
#include <QtGui/QGridLayout>
#include <QtGui/QGroupBox>
#include <QtGui/QHeaderView>
#include <QtGui/QLabel>

QT_BEGIN_NAMESPACE

class Ui_AnalysisInputDialog
{
public:
    QGridLayout *gridLayout_2;
    QGroupBox *groupBoxTopoInputs;
    QFormLayout *formLayout;
    QLabel *labelInterpMethod;
    QComboBox *comboBoxInterpMethod;
    QLabel *labelProfileSpacing;
    QDoubleSpinBox *doubleSpinBoxProfileSpacing;
    QLabel *labelDatumCode;
    QComboBox *comboBoxDatumCode;
    QGroupBox *groupBoxTerrainInputs;
    QGridLayout *gridLayout;
    QLabel *labelHumidity;
    QDoubleSpinBox *doubleSpinBoxHumidity;
    QLabel *labelRefractivity;
    QDoubleSpinBox *doubleSpinBoxRefractivity;
    QLabel *labelGroundType;
    QComboBox *comboBoxGroundType;
    QLabel *labelPermittivity;
    QDoubleSpinBox *doubleSpinBoxPermittivity;
    QLabel *labelConductivity;
    QDoubleSpinBox *doubleSpinBoxConductivity;
    QDialogButtonBox *buttonBox;

    void setupUi(QDialog *AnalysisInputDialog)
    {
        if (AnalysisInputDialog->objectName().isEmpty())
            AnalysisInputDialog->setObjectName(QString::fromUtf8("AnalysisInputDialog"));
        AnalysisInputDialog->resize(488, 389);
        AnalysisInputDialog->setModal(true);
        gridLayout_2 = new QGridLayout(AnalysisInputDialog);
        gridLayout_2->setObjectName(QString::fromUtf8("gridLayout_2"));
        groupBoxTopoInputs = new QGroupBox(AnalysisInputDialog);
        groupBoxTopoInputs->setObjectName(QString::fromUtf8("groupBoxTopoInputs"));
        formLayout = new QFormLayout(groupBoxTopoInputs);
        formLayout->setObjectName(QString::fromUtf8("formLayout"));
        labelInterpMethod = new QLabel(groupBoxTopoInputs);
        labelInterpMethod->setObjectName(QString::fromUtf8("labelInterpMethod"));

        formLayout->setWidget(0, QFormLayout::LabelRole, labelInterpMethod);

        comboBoxInterpMethod = new QComboBox(groupBoxTopoInputs);
        comboBoxInterpMethod->setObjectName(QString::fromUtf8("comboBoxInterpMethod"));

        formLayout->setWidget(0, QFormLayout::FieldRole, comboBoxInterpMethod);

        labelProfileSpacing = new QLabel(groupBoxTopoInputs);
        labelProfileSpacing->setObjectName(QString::fromUtf8("labelProfileSpacing"));

        formLayout->setWidget(1, QFormLayout::LabelRole, labelProfileSpacing);

        doubleSpinBoxProfileSpacing = new QDoubleSpinBox(groupBoxTopoInputs);
        doubleSpinBoxProfileSpacing->setObjectName(QString::fromUtf8("doubleSpinBoxProfileSpacing"));
        doubleSpinBoxProfileSpacing->setDecimals(0);
        doubleSpinBoxProfileSpacing->setMinimum(3);
        doubleSpinBoxProfileSpacing->setMaximum(30);
        doubleSpinBoxProfileSpacing->setSingleStep(3);

        formLayout->setWidget(1, QFormLayout::FieldRole, doubleSpinBoxProfileSpacing);

        labelDatumCode = new QLabel(groupBoxTopoInputs);
        labelDatumCode->setObjectName(QString::fromUtf8("labelDatumCode"));

        formLayout->setWidget(2, QFormLayout::LabelRole, labelDatumCode);

        comboBoxDatumCode = new QComboBox(groupBoxTopoInputs);
        comboBoxDatumCode->setObjectName(QString::fromUtf8("comboBoxDatumCode"));

        formLayout->setWidget(2, QFormLayout::FieldRole, comboBoxDatumCode);


        gridLayout_2->addWidget(groupBoxTopoInputs, 0, 0, 1, 1);

        groupBoxTerrainInputs = new QGroupBox(AnalysisInputDialog);
        groupBoxTerrainInputs->setObjectName(QString::fromUtf8("groupBoxTerrainInputs"));
        gridLayout = new QGridLayout(groupBoxTerrainInputs);
        gridLayout->setObjectName(QString::fromUtf8("gridLayout"));
        labelHumidity = new QLabel(groupBoxTerrainInputs);
        labelHumidity->setObjectName(QString::fromUtf8("labelHumidity"));

        gridLayout->addWidget(labelHumidity, 0, 0, 1, 1);

        doubleSpinBoxHumidity = new QDoubleSpinBox(groupBoxTerrainInputs);
        doubleSpinBoxHumidity->setObjectName(QString::fromUtf8("doubleSpinBoxHumidity"));
        doubleSpinBoxHumidity->setMaximum(1000);
        doubleSpinBoxHumidity->setValue(10);

        gridLayout->addWidget(doubleSpinBoxHumidity, 0, 1, 1, 1);

        labelRefractivity = new QLabel(groupBoxTerrainInputs);
        labelRefractivity->setObjectName(QString::fromUtf8("labelRefractivity"));

        gridLayout->addWidget(labelRefractivity, 1, 0, 1, 1);

        doubleSpinBoxRefractivity = new QDoubleSpinBox(groupBoxTerrainInputs);
        doubleSpinBoxRefractivity->setObjectName(QString::fromUtf8("doubleSpinBoxRefractivity"));
        doubleSpinBoxRefractivity->setMinimum(-1000);
        doubleSpinBoxRefractivity->setMaximum(10000);
        doubleSpinBoxRefractivity->setValue(301);

        gridLayout->addWidget(doubleSpinBoxRefractivity, 1, 1, 1, 1);

        labelGroundType = new QLabel(groupBoxTerrainInputs);
        labelGroundType->setObjectName(QString::fromUtf8("labelGroundType"));

        gridLayout->addWidget(labelGroundType, 2, 0, 1, 1);

        comboBoxGroundType = new QComboBox(groupBoxTerrainInputs);
        comboBoxGroundType->setObjectName(QString::fromUtf8("comboBoxGroundType"));

        gridLayout->addWidget(comboBoxGroundType, 2, 1, 1, 1);

        labelPermittivity = new QLabel(groupBoxTerrainInputs);
        labelPermittivity->setObjectName(QString::fromUtf8("labelPermittivity"));

        gridLayout->addWidget(labelPermittivity, 3, 0, 1, 1);

        doubleSpinBoxPermittivity = new QDoubleSpinBox(groupBoxTerrainInputs);
        doubleSpinBoxPermittivity->setObjectName(QString::fromUtf8("doubleSpinBoxPermittivity"));
        doubleSpinBoxPermittivity->setDecimals(6);
        doubleSpinBoxPermittivity->setMaximum(1000);
        doubleSpinBoxPermittivity->setValue(15);

        gridLayout->addWidget(doubleSpinBoxPermittivity, 3, 1, 1, 1);

        labelConductivity = new QLabel(groupBoxTerrainInputs);
        labelConductivity->setObjectName(QString::fromUtf8("labelConductivity"));

        gridLayout->addWidget(labelConductivity, 4, 0, 1, 1);

        doubleSpinBoxConductivity = new QDoubleSpinBox(groupBoxTerrainInputs);
        doubleSpinBoxConductivity->setObjectName(QString::fromUtf8("doubleSpinBoxConductivity"));
        doubleSpinBoxConductivity->setDecimals(6);
        doubleSpinBoxConductivity->setMaximum(10);
        doubleSpinBoxConductivity->setSingleStep(0.01);
        doubleSpinBoxConductivity->setValue(0.03);

        gridLayout->addWidget(doubleSpinBoxConductivity, 4, 1, 1, 1);


        gridLayout_2->addWidget(groupBoxTerrainInputs, 1, 0, 1, 1);

        buttonBox = new QDialogButtonBox(AnalysisInputDialog);
        buttonBox->setObjectName(QString::fromUtf8("buttonBox"));
        buttonBox->setOrientation(Qt::Horizontal);
        buttonBox->setStandardButtons(QDialogButtonBox::Cancel|QDialogButtonBox::Ok);

        gridLayout_2->addWidget(buttonBox, 2, 0, 1, 1);


        retranslateUi(AnalysisInputDialog);
        QObject::connect(buttonBox, SIGNAL(accepted()), AnalysisInputDialog, SLOT(accept()));
        QObject::connect(buttonBox, SIGNAL(rejected()), AnalysisInputDialog, SLOT(reject()));

        comboBoxInterpMethod->setCurrentIndex(2);


        QMetaObject::connectSlotsByName(AnalysisInputDialog);
    } // setupUi

    void retranslateUi(QDialog *AnalysisInputDialog)
    {
        AnalysisInputDialog->setWindowTitle(QApplication::translate("AnalysisInputDialog", "Analysis Input", 0, QApplication::UnicodeUTF8));
        groupBoxTopoInputs->setTitle(QApplication::translate("AnalysisInputDialog", "Topographic Extraction Inputs", 0, QApplication::UnicodeUTF8));
        labelInterpMethod->setText(QApplication::translate("AnalysisInputDialog", "Interpolation Method", 0, QApplication::UnicodeUTF8));
        comboBoxInterpMethod->clear();
        comboBoxInterpMethod->insertItems(0, QStringList()
         << QApplication::translate("AnalysisInputDialog", "Nearest", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Highest", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Four-Point", 0, QApplication::UnicodeUTF8)
        );
        labelProfileSpacing->setText(QApplication::translate("AnalysisInputDialog", "Profile Spacing", 0, QApplication::UnicodeUTF8));
        doubleSpinBoxProfileSpacing->setSpecialValueText(QApplication::translate("AnalysisInputDialog", "Use Topo File Spacing", 0, QApplication::UnicodeUTF8));
        doubleSpinBoxProfileSpacing->setSuffix(QApplication::translate("AnalysisInputDialog", " seconds", 0, QApplication::UnicodeUTF8));
        labelDatumCode->setText(QApplication::translate("AnalysisInputDialog", "Datum Code", 0, QApplication::UnicodeUTF8));
        comboBoxDatumCode->clear();
        comboBoxDatumCode->insertItems(0, QStringList()
         << QApplication::translate("AnalysisInputDialog", "World Geodetic System 1984", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "North American Datum NAD 27", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "European", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Tokyo", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Great Britain", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Maui (Old Hawaiian)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Oahu (Old Hawaiian)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Kauai (Old Hawaiian)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Kwajalein Atoll (Wake-Eniwetok)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Wake Island (Wake-Eniwetok)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Eniwetok Atoll (Wake-Eniwetok)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Wake Island Astro 1952", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Guam 1963", 0, QApplication::UnicodeUTF8)
        );
        groupBoxTerrainInputs->setTitle(QApplication::translate("AnalysisInputDialog", "Terrain Propagation Inputs", 0, QApplication::UnicodeUTF8));
        labelHumidity->setText(QApplication::translate("AnalysisInputDialog", "Humidity", 0, QApplication::UnicodeUTF8));
        doubleSpinBoxHumidity->setSuffix(QApplication::translate("AnalysisInputDialog", " g/m3", 0, QApplication::UnicodeUTF8));
        labelRefractivity->setText(QApplication::translate("AnalysisInputDialog", "Sea-Level Atmospheric Refractivity", 0, QApplication::UnicodeUTF8));
        doubleSpinBoxRefractivity->setSuffix(QApplication::translate("AnalysisInputDialog", " M-units", 0, QApplication::UnicodeUTF8));
        labelGroundType->setText(QApplication::translate("AnalysisInputDialog", "CCIR Ground Type", 0, QApplication::UnicodeUTF8));
        comboBoxGroundType->clear();
        comboBoxGroundType->insertItems(0, QStringList()
         << QApplication::translate("AnalysisInputDialog", "None", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Sea Water (20 Degrees C)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Medium Dry Ground", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Very Dry Ground", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Pure Water (Not Used)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Ice (Fresh Water, -1 Degree C)", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("AnalysisInputDialog", "Ice (Fresh Water, -10 Degree C)", 0, QApplication::UnicodeUTF8)
        );
        labelPermittivity->setText(QApplication::translate("AnalysisInputDialog", "Relative Permittivity of Earth Surface", 0, QApplication::UnicodeUTF8));
        labelConductivity->setText(QApplication::translate("AnalysisInputDialog", "Conductivity of Earth Surface", 0, QApplication::UnicodeUTF8));
        doubleSpinBoxConductivity->setSuffix(QApplication::translate("AnalysisInputDialog", " S/m", 0, QApplication::UnicodeUTF8));
    } // retranslateUi

};

namespace Ui {
    class AnalysisInputDialog: public Ui_AnalysisInputDialog {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_ANALYSISINPUTDIALOG_H
