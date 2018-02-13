/////////////////////////////////////////////////////////////////////////////
//
// COPYRIGHT: Copyright 2006
//     Alion Science and Technology
//     US Govt Retains rights in accordance
//     with DoD FAR Supp 252.227 - 7013.
//
/////////////////////////////////////////////////////////////////////////////


#if ! defined( DISTRIBUTIONDIALOG_H )
#define DISTRIBUTIONDIALOG_H


#include <QtGui/QDialog>
#include "ui_DistributionDialog.h"


class DistributionDialog : public QDialog, private Ui::DistributionDialog
{

public:

    DistributionDialog( QWidget * parent = 0 );

private:

};


#endif // DISTRIBUTIONDIALOG_H
