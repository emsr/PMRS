

#include <QtGui/QDialog>


#include "ui_AnalysisInputDialog.h"


class AnalysisInputDialog : public QDialog, private Ui::AnalysisInputDialog
{

    Q_OBJECT

public:

    AnalysisInputDialog( QWidget * parent = 0 );

public slots:

    void accept( void );

private slots:

    void setGroundParams( int index );

private:

    static const double permittivity[9];
    static const double conductivity[9];

};
