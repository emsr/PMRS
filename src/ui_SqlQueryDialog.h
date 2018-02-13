/********************************************************************************
** Form generated from reading UI file 'SqlQueryDialog.ui'
**
** Created by: Qt User Interface Compiler version 4.8.7
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_SQLQUERYDIALOG_H
#define UI_SQLQUERYDIALOG_H

#include <QtCore/QVariant>
#include <QtGui/QAction>
#include <QtGui/QApplication>
#include <QtGui/QButtonGroup>
#include <QtGui/QComboBox>
#include <QtGui/QDialog>
#include <QtGui/QDialogButtonBox>
#include <QtGui/QDoubleSpinBox>
#include <QtGui/QGridLayout>
#include <QtGui/QHeaderView>
#include <QtGui/QLabel>
#include <QtGui/QListView>
#include "SqlSelectOperator.h"

QT_BEGIN_NAMESPACE

class Ui_SqlQueryDialog
{
public:
    QGridLayout *gridLayout;
    QLabel *labelLocation;
    QListView *listViewLocation;
    QLabel *labelFrequency;
    SqlSelectOperator *comboBoxFrequency;
    QDoubleSpinBox *doubleSpinBoxFreqBegin;
    QDoubleSpinBox *doubleSpinBoxFreqEnd;
    QLabel *labelLinkDist;
    SqlSelectOperator *comboBoxLinkDist;
    QDoubleSpinBox *doubleSpinBoxLinkDistBegin;
    QDoubleSpinBox *doubleSpinBoxLinkDistEnd;
    QLabel *labelTxAntHt;
    SqlSelectOperator *comboBoxTxAntHt;
    QDoubleSpinBox *doubleSpinBoxTxAntHtBegin;
    QDoubleSpinBox *doubleSpinBoxTxAntHtEnd;
    QLabel *labelRxAntHt;
    SqlSelectOperator *comboBoxRxAntHt;
    QDoubleSpinBox *doubleSpinBoxRxAntHtBegin;
    QDoubleSpinBox *doubleSpinBoxRxAntHtEnd;
    QLabel *labelPolar;
    QComboBox *comboBoxPolar;
    QDialogButtonBox *buttonBox;

    void setupUi(QDialog *SqlQueryDialog)
    {
        if (SqlQueryDialog->objectName().isEmpty())
            SqlQueryDialog->setObjectName(QString::fromUtf8("SqlQueryDialog"));
        SqlQueryDialog->resize(530, 392);
        SqlQueryDialog->setModal(true);
        gridLayout = new QGridLayout(SqlQueryDialog);
        gridLayout->setObjectName(QString::fromUtf8("gridLayout"));
        labelLocation = new QLabel(SqlQueryDialog);
        labelLocation->setObjectName(QString::fromUtf8("labelLocation"));

        gridLayout->addWidget(labelLocation, 0, 0, 1, 1);

        listViewLocation = new QListView(SqlQueryDialog);
        listViewLocation->setObjectName(QString::fromUtf8("listViewLocation"));
        listViewLocation->setSelectionMode(QAbstractItemView::ExtendedSelection);

        gridLayout->addWidget(listViewLocation, 0, 1, 1, 3);

        labelFrequency = new QLabel(SqlQueryDialog);
        labelFrequency->setObjectName(QString::fromUtf8("labelFrequency"));

        gridLayout->addWidget(labelFrequency, 1, 0, 1, 1);

        comboBoxFrequency = new SqlSelectOperator(SqlQueryDialog);
        comboBoxFrequency->setObjectName(QString::fromUtf8("comboBoxFrequency"));

        gridLayout->addWidget(comboBoxFrequency, 1, 1, 1, 1);

        doubleSpinBoxFreqBegin = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxFreqBegin->setObjectName(QString::fromUtf8("doubleSpinBoxFreqBegin"));
        doubleSpinBoxFreqBegin->setMaximum(1e+06);

        gridLayout->addWidget(doubleSpinBoxFreqBegin, 1, 2, 1, 1);

        doubleSpinBoxFreqEnd = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxFreqEnd->setObjectName(QString::fromUtf8("doubleSpinBoxFreqEnd"));
        doubleSpinBoxFreqEnd->setMaximum(1e+06);

        gridLayout->addWidget(doubleSpinBoxFreqEnd, 1, 3, 1, 1);

        labelLinkDist = new QLabel(SqlQueryDialog);
        labelLinkDist->setObjectName(QString::fromUtf8("labelLinkDist"));

        gridLayout->addWidget(labelLinkDist, 2, 0, 1, 1);

        comboBoxLinkDist = new SqlSelectOperator(SqlQueryDialog);
        comboBoxLinkDist->setObjectName(QString::fromUtf8("comboBoxLinkDist"));

        gridLayout->addWidget(comboBoxLinkDist, 2, 1, 1, 1);

        doubleSpinBoxLinkDistBegin = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxLinkDistBegin->setObjectName(QString::fromUtf8("doubleSpinBoxLinkDistBegin"));
        doubleSpinBoxLinkDistBegin->setMaximum(1000);

        gridLayout->addWidget(doubleSpinBoxLinkDistBegin, 2, 2, 1, 1);

        doubleSpinBoxLinkDistEnd = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxLinkDistEnd->setObjectName(QString::fromUtf8("doubleSpinBoxLinkDistEnd"));
        doubleSpinBoxLinkDistEnd->setMaximum(1000);

        gridLayout->addWidget(doubleSpinBoxLinkDistEnd, 2, 3, 1, 1);

        labelTxAntHt = new QLabel(SqlQueryDialog);
        labelTxAntHt->setObjectName(QString::fromUtf8("labelTxAntHt"));

        gridLayout->addWidget(labelTxAntHt, 3, 0, 1, 1);

        comboBoxTxAntHt = new SqlSelectOperator(SqlQueryDialog);
        comboBoxTxAntHt->setObjectName(QString::fromUtf8("comboBoxTxAntHt"));

        gridLayout->addWidget(comboBoxTxAntHt, 3, 1, 1, 1);

        doubleSpinBoxTxAntHtBegin = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxTxAntHtBegin->setObjectName(QString::fromUtf8("doubleSpinBoxTxAntHtBegin"));
        doubleSpinBoxTxAntHtBegin->setMaximum(40000);

        gridLayout->addWidget(doubleSpinBoxTxAntHtBegin, 3, 2, 1, 1);

        doubleSpinBoxTxAntHtEnd = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxTxAntHtEnd->setObjectName(QString::fromUtf8("doubleSpinBoxTxAntHtEnd"));
        doubleSpinBoxTxAntHtEnd->setMaximum(40000);

        gridLayout->addWidget(doubleSpinBoxTxAntHtEnd, 3, 3, 1, 1);

        labelRxAntHt = new QLabel(SqlQueryDialog);
        labelRxAntHt->setObjectName(QString::fromUtf8("labelRxAntHt"));

        gridLayout->addWidget(labelRxAntHt, 4, 0, 1, 1);

        comboBoxRxAntHt = new SqlSelectOperator(SqlQueryDialog);
        comboBoxRxAntHt->setObjectName(QString::fromUtf8("comboBoxRxAntHt"));

        gridLayout->addWidget(comboBoxRxAntHt, 4, 1, 1, 1);

        doubleSpinBoxRxAntHtBegin = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxRxAntHtBegin->setObjectName(QString::fromUtf8("doubleSpinBoxRxAntHtBegin"));
        doubleSpinBoxRxAntHtBegin->setMaximum(40000);

        gridLayout->addWidget(doubleSpinBoxRxAntHtBegin, 4, 2, 1, 1);

        doubleSpinBoxRxAntHtEnd = new QDoubleSpinBox(SqlQueryDialog);
        doubleSpinBoxRxAntHtEnd->setObjectName(QString::fromUtf8("doubleSpinBoxRxAntHtEnd"));
        doubleSpinBoxRxAntHtEnd->setMaximum(40000);

        gridLayout->addWidget(doubleSpinBoxRxAntHtEnd, 4, 3, 1, 1);

        labelPolar = new QLabel(SqlQueryDialog);
        labelPolar->setObjectName(QString::fromUtf8("labelPolar"));

        gridLayout->addWidget(labelPolar, 5, 0, 1, 1);

        comboBoxPolar = new QComboBox(SqlQueryDialog);
        comboBoxPolar->setObjectName(QString::fromUtf8("comboBoxPolar"));

        gridLayout->addWidget(comboBoxPolar, 5, 1, 1, 1);

        buttonBox = new QDialogButtonBox(SqlQueryDialog);
        buttonBox->setObjectName(QString::fromUtf8("buttonBox"));
        buttonBox->setOrientation(Qt::Horizontal);
        buttonBox->setStandardButtons(QDialogButtonBox::Cancel|QDialogButtonBox::Ok);

        gridLayout->addWidget(buttonBox, 6, 0, 1, 4);

#ifndef QT_NO_SHORTCUT
        labelLocation->setBuddy(listViewLocation);
        labelFrequency->setBuddy(comboBoxFrequency);
        labelLinkDist->setBuddy(comboBoxLinkDist);
        labelTxAntHt->setBuddy(comboBoxTxAntHt);
        labelRxAntHt->setBuddy(comboBoxRxAntHt);
        labelPolar->setBuddy(comboBoxPolar);
#endif // QT_NO_SHORTCUT
        QWidget::setTabOrder(listViewLocation, comboBoxFrequency);
        QWidget::setTabOrder(comboBoxFrequency, doubleSpinBoxFreqBegin);
        QWidget::setTabOrder(doubleSpinBoxFreqBegin, doubleSpinBoxFreqEnd);
        QWidget::setTabOrder(doubleSpinBoxFreqEnd, comboBoxLinkDist);
        QWidget::setTabOrder(comboBoxLinkDist, doubleSpinBoxLinkDistBegin);
        QWidget::setTabOrder(doubleSpinBoxLinkDistBegin, doubleSpinBoxLinkDistEnd);
        QWidget::setTabOrder(doubleSpinBoxLinkDistEnd, comboBoxTxAntHt);
        QWidget::setTabOrder(comboBoxTxAntHt, doubleSpinBoxTxAntHtBegin);
        QWidget::setTabOrder(doubleSpinBoxTxAntHtBegin, doubleSpinBoxTxAntHtEnd);
        QWidget::setTabOrder(doubleSpinBoxTxAntHtEnd, comboBoxRxAntHt);
        QWidget::setTabOrder(comboBoxRxAntHt, doubleSpinBoxRxAntHtBegin);
        QWidget::setTabOrder(doubleSpinBoxRxAntHtBegin, doubleSpinBoxRxAntHtEnd);
        QWidget::setTabOrder(doubleSpinBoxRxAntHtEnd, comboBoxPolar);
        QWidget::setTabOrder(comboBoxPolar, buttonBox);

        retranslateUi(SqlQueryDialog);
        QObject::connect(buttonBox, SIGNAL(accepted()), SqlQueryDialog, SLOT(accept()));
        QObject::connect(buttonBox, SIGNAL(rejected()), SqlQueryDialog, SLOT(reject()));

        QMetaObject::connectSlotsByName(SqlQueryDialog);
    } // setupUi

    void retranslateUi(QDialog *SqlQueryDialog)
    {
        SqlQueryDialog->setWindowTitle(QApplication::translate("SqlQueryDialog", "Query Select Criteria", 0, QApplication::UnicodeUTF8));
        labelLocation->setText(QApplication::translate("SqlQueryDialog", "Location", 0, QApplication::UnicodeUTF8));
        labelFrequency->setText(QApplication::translate("SqlQueryDialog", "Frequency", 0, QApplication::UnicodeUTF8));
        labelLinkDist->setText(QApplication::translate("SqlQueryDialog", "Link Distance", 0, QApplication::UnicodeUTF8));
        labelTxAntHt->setText(QApplication::translate("SqlQueryDialog", "Tx Antenna Height", 0, QApplication::UnicodeUTF8));
        labelRxAntHt->setText(QApplication::translate("SqlQueryDialog", "Rx Antenna Height", 0, QApplication::UnicodeUTF8));
        labelPolar->setText(QApplication::translate("SqlQueryDialog", "Polarization", 0, QApplication::UnicodeUTF8));
        comboBoxPolar->clear();
        comboBoxPolar->insertItems(0, QStringList()
         << QApplication::translate("SqlQueryDialog", "Like Polarization", 0, QApplication::UnicodeUTF8)
         << QApplication::translate("SqlQueryDialog", "Cross Polarization", 0, QApplication::UnicodeUTF8)
        );
    } // retranslateUi

};

namespace Ui {
    class SqlQueryDialog: public Ui_SqlQueryDialog {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_SQLQUERYDIALOG_H
