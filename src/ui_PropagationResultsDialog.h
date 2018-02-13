/********************************************************************************
** Form generated from reading UI file 'PropagationResultsDialog.ui'
**
** Created by: Qt User Interface Compiler version 4.8.7
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_PROPAGATIONRESULTSDIALOG_H
#define UI_PROPAGATIONRESULTSDIALOG_H

#include <QtCore/QVariant>
#include <QtGui/QAction>
#include <QtGui/QApplication>
#include <QtGui/QButtonGroup>
#include <QtGui/QDialog>
#include <QtGui/QDialogButtonBox>
#include <QtGui/QGridLayout>
#include <QtGui/QHeaderView>
#include <QtGui/QTableView>

QT_BEGIN_NAMESPACE

class Ui_PropagationResultsDialog
{
public:
    QGridLayout *gridLayout;
    QTableView *tableView;
    QDialogButtonBox *buttonBox;

    void setupUi(QDialog *PropagationResultsDialog)
    {
        if (PropagationResultsDialog->objectName().isEmpty())
            PropagationResultsDialog->setObjectName(QString::fromUtf8("PropagationResultsDialog"));
        PropagationResultsDialog->resize(695, 500);
        PropagationResultsDialog->setModal(true);
        gridLayout = new QGridLayout(PropagationResultsDialog);
        gridLayout->setObjectName(QString::fromUtf8("gridLayout"));
        tableView = new QTableView(PropagationResultsDialog);
        tableView->setObjectName(QString::fromUtf8("tableView"));

        gridLayout->addWidget(tableView, 0, 0, 1, 1);

        buttonBox = new QDialogButtonBox(PropagationResultsDialog);
        buttonBox->setObjectName(QString::fromUtf8("buttonBox"));
        buttonBox->setOrientation(Qt::Horizontal);
        buttonBox->setStandardButtons(QDialogButtonBox::Cancel|QDialogButtonBox::Ok);

        gridLayout->addWidget(buttonBox, 1, 0, 1, 1);


        retranslateUi(PropagationResultsDialog);
        QObject::connect(buttonBox, SIGNAL(accepted()), PropagationResultsDialog, SLOT(accept()));
        QObject::connect(buttonBox, SIGNAL(rejected()), PropagationResultsDialog, SLOT(reject()));

        QMetaObject::connectSlotsByName(PropagationResultsDialog);
    } // setupUi

    void retranslateUi(QDialog *PropagationResultsDialog)
    {
        PropagationResultsDialog->setWindowTitle(QApplication::translate("PropagationResultsDialog", "Propagation Results", 0, QApplication::UnicodeUTF8));
    } // retranslateUi

};

namespace Ui {
    class PropagationResultsDialog: public Ui_PropagationResultsDialog {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_PROPAGATIONRESULTSDIALOG_H
