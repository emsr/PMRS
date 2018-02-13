/********************************************************************************
** Form generated from reading UI file 'SqlQueryResultsDialog.ui'
**
** Created by: Qt User Interface Compiler version 4.8.7
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_SQLQUERYRESULTSDIALOG_H
#define UI_SQLQUERYRESULTSDIALOG_H

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

class Ui_SqlQueryResultsDialog
{
public:
    QGridLayout *gridLayout;
    QTableView *tableViewQueryResults;
    QDialogButtonBox *buttonBox;

    void setupUi(QDialog *SqlQueryResultsDialog)
    {
        if (SqlQueryResultsDialog->objectName().isEmpty())
            SqlQueryResultsDialog->setObjectName(QString::fromUtf8("SqlQueryResultsDialog"));
        SqlQueryResultsDialog->resize(558, 439);
        SqlQueryResultsDialog->setModal(true);
        gridLayout = new QGridLayout(SqlQueryResultsDialog);
        gridLayout->setObjectName(QString::fromUtf8("gridLayout"));
        tableViewQueryResults = new QTableView(SqlQueryResultsDialog);
        tableViewQueryResults->setObjectName(QString::fromUtf8("tableViewQueryResults"));

        gridLayout->addWidget(tableViewQueryResults, 0, 0, 1, 1);

        buttonBox = new QDialogButtonBox(SqlQueryResultsDialog);
        buttonBox->setObjectName(QString::fromUtf8("buttonBox"));
        buttonBox->setOrientation(Qt::Horizontal);
        buttonBox->setStandardButtons(QDialogButtonBox::Cancel|QDialogButtonBox::Ok);

        gridLayout->addWidget(buttonBox, 1, 0, 1, 1);


        retranslateUi(SqlQueryResultsDialog);
        QObject::connect(buttonBox, SIGNAL(accepted()), SqlQueryResultsDialog, SLOT(accept()));
        QObject::connect(buttonBox, SIGNAL(rejected()), SqlQueryResultsDialog, SLOT(reject()));

        QMetaObject::connectSlotsByName(SqlQueryResultsDialog);
    } // setupUi

    void retranslateUi(QDialog *SqlQueryResultsDialog)
    {
        SqlQueryResultsDialog->setWindowTitle(QApplication::translate("SqlQueryResultsDialog", "Query Results", 0, QApplication::UnicodeUTF8));
    } // retranslateUi

};

namespace Ui {
    class SqlQueryResultsDialog: public Ui_SqlQueryResultsDialog {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_SQLQUERYRESULTSDIALOG_H
