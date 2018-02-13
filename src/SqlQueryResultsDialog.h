

#include <QtGui/QDialog>


#include "ui_SqlQueryResultsDialog.h"


class SqlQueryResultsDialog : public QDialog, private Ui::SqlQueryResultsDialog
{

    Q_OBJECT

public:

    SqlQueryResultsDialog( QWidget * parent = 0 );

public slots:

    void accept( void );

};
