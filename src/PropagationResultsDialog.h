

#include <QtGui/QDialog>


#include "ui_PropagationResultsDialog.h"


class PropagationResultsDialog : public QDialog, private Ui::PropagationResultsDialog
{

    Q_OBJECT

public:

    PropagationResultsDialog( QWidget * parent = 0 );

public slots:

    void accept( void );

};
