/********************************************************************************
** Form generated from reading UI file 'DistributionDialog.ui'
**
** Created by: Qt User Interface Compiler version 4.8.7
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_DISTRIBUTIONDIALOG_H
#define UI_DISTRIBUTIONDIALOG_H

#include <QtCore/QVariant>
#include <QtGui/QAction>
#include <QtGui/QApplication>
#include <QtGui/QButtonGroup>
#include <QtGui/QDialog>
#include <QtGui/QHBoxLayout>
#include <QtGui/QHeaderView>
#include <QtGui/QLabel>
#include <QtGui/QPushButton>
#include <QtGui/QSpacerItem>
#include <QtGui/QVBoxLayout>

QT_BEGIN_NAMESPACE

class Ui_DistributionDialog
{
public:
    QVBoxLayout *vboxLayout;
    QLabel *labelBlank;
    QLabel *labelWARNING;
    QLabel *labelDESTRUCTION;
    QLabel *labelDISTRIBUTION;
    QHBoxLayout *hboxLayout;
    QSpacerItem *spacerItem;
    QPushButton *okButton;
    QSpacerItem *spacerItem1;

    void setupUi(QDialog *DistributionDialog)
    {
        if (DistributionDialog->objectName().isEmpty())
            DistributionDialog->setObjectName(QString::fromUtf8("DistributionDialog"));
        DistributionDialog->resize(492, 290);
        vboxLayout = new QVBoxLayout(DistributionDialog);
#ifndef Q_OS_MAC
        vboxLayout->setSpacing(6);
#endif
#ifndef Q_OS_MAC
        vboxLayout->setContentsMargins(9, 9, 9, 9);
#endif
        vboxLayout->setObjectName(QString::fromUtf8("vboxLayout"));
        labelBlank = new QLabel(DistributionDialog);
        labelBlank->setObjectName(QString::fromUtf8("labelBlank"));
        QSizePolicy sizePolicy(static_cast<QSizePolicy::Policy>(5), static_cast<QSizePolicy::Policy>(0));
        sizePolicy.setHorizontalStretch(0);
        sizePolicy.setVerticalStretch(0);
        sizePolicy.setHeightForWidth(labelBlank->sizePolicy().hasHeightForWidth());
        labelBlank->setSizePolicy(sizePolicy);

        vboxLayout->addWidget(labelBlank);

        labelWARNING = new QLabel(DistributionDialog);
        labelWARNING->setObjectName(QString::fromUtf8("labelWARNING"));
        QSizePolicy sizePolicy1(static_cast<QSizePolicy::Policy>(5), static_cast<QSizePolicy::Policy>(5));
        sizePolicy1.setHorizontalStretch(0);
        sizePolicy1.setVerticalStretch(0);
        sizePolicy1.setHeightForWidth(labelWARNING->sizePolicy().hasHeightForWidth());
        labelWARNING->setSizePolicy(sizePolicy1);
        labelWARNING->setAlignment(Qt::AlignCenter);
        labelWARNING->setWordWrap(true);

        vboxLayout->addWidget(labelWARNING);

        labelDESTRUCTION = new QLabel(DistributionDialog);
        labelDESTRUCTION->setObjectName(QString::fromUtf8("labelDESTRUCTION"));
        labelDESTRUCTION->setAlignment(Qt::AlignCenter);
        labelDESTRUCTION->setWordWrap(true);

        vboxLayout->addWidget(labelDESTRUCTION);

        labelDISTRIBUTION = new QLabel(DistributionDialog);
        labelDISTRIBUTION->setObjectName(QString::fromUtf8("labelDISTRIBUTION"));
        labelDISTRIBUTION->setAlignment(Qt::AlignCenter);
        labelDISTRIBUTION->setWordWrap(true);

        vboxLayout->addWidget(labelDISTRIBUTION);

        hboxLayout = new QHBoxLayout();
#ifndef Q_OS_MAC
        hboxLayout->setSpacing(6);
#endif
        hboxLayout->setContentsMargins(0, 0, 0, 0);
        hboxLayout->setObjectName(QString::fromUtf8("hboxLayout"));
        spacerItem = new QSpacerItem(151, 38, QSizePolicy::Expanding, QSizePolicy::Minimum);

        hboxLayout->addItem(spacerItem);

        okButton = new QPushButton(DistributionDialog);
        okButton->setObjectName(QString::fromUtf8("okButton"));

        hboxLayout->addWidget(okButton);

        spacerItem1 = new QSpacerItem(40, 20, QSizePolicy::Expanding, QSizePolicy::Minimum);

        hboxLayout->addItem(spacerItem1);


        vboxLayout->addLayout(hboxLayout);


        retranslateUi(DistributionDialog);
        QObject::connect(okButton, SIGNAL(clicked()), DistributionDialog, SLOT(accept()));

        QMetaObject::connectSlotsByName(DistributionDialog);
    } // setupUi

    void retranslateUi(QDialog *DistributionDialog)
    {
        DistributionDialog->setWindowTitle(QApplication::translate("DistributionDialog", "Export Control Notice", 0, QApplication::UnicodeUTF8));
        labelBlank->setText(QString());
        labelWARNING->setText(QApplication::translate("DistributionDialog", "WARNING - This software contains technical data whose export is restricted by the Arms Export Control Act (Title 22, U.S.C., Sec 2751 et. seq.), or the Export Administration Act of 1979, as amended (Title 50, U.S.C., App 2401 et. seq.), or Executive Order 12924.  Violations of these export laws are subject to severe criminal penalties.  Disseminate in accordance with provisions of DoD Directive 5230.25.", 0, QApplication::UnicodeUTF8));
        labelDESTRUCTION->setText(QApplication::translate("DistributionDialog", "DESTRUCTION NOTICE - For unclassified limited material, destroy by any method that will prevent disclosure of contents or reconstruction of the material.", 0, QApplication::UnicodeUTF8));
        labelDISTRIBUTION->setText(QApplication::translate("DistributionDialog", "DISTRIBUTION NOTICE - Further dissemination only as directed by DoD JSC, (Feb. 2002), or higher DoD authority.", 0, QApplication::UnicodeUTF8));
        okButton->setText(QApplication::translate("DistributionDialog", "OK", 0, QApplication::UnicodeUTF8));
    } // retranslateUi

};

namespace Ui {
    class DistributionDialog: public Ui_DistributionDialog {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_DISTRIBUTIONDIALOG_H
