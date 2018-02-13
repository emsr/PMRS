#include <QtGui/QApplication>
#include <QtGui/QSplashScreen>
#include <QtCore/QTimer>


#include "DistributionDialog.h"
#include "SqlQueryDialog.h"
#include "SqlQueryResultsDialog.h"
#include "AnalysisInputDialog.h"
#include "PropagationResultsDialog.h"


int
main( int n_app_args, char **app_arg )
{
    QApplication application( n_app_args, app_arg );    

    //  Put up the splash screen.
    QPixmap pixmap( ":/artwork/splash.png" );
    QSplashScreen splash( pixmap );
    splash.show();

    //  Set a timer on the splash screen (esp. in case something pops under it).
    QTimer::singleShot( 10000, &splash, SLOT(close()) );

    DistributionDialog distribution( &splash );
    distribution.exec();

    //  If the splash screen lasts this long close it.
    splash.finish( &distribution );

    //  Last one out turn off the lights.
    application.connect( &application, SIGNAL(lastWindowClosed()),
                         &application, SLOT(quit()) );

    //int result = QDialog::Rejected;

    SqlQueryDialog * sqlQueryDialog = new SqlQueryDialog;
    SqlQueryResultsDialog * sqlQueryResultsDialog = new SqlQueryResultsDialog;
    AnalysisInputDialog * analysisInputDialog = new AnalysisInputDialog;
    PropagationResultsDialog * propagationResultsDialog = new PropagationResultsDialog;

    QObject::connect( sqlQueryDialog, SIGNAL(accepted()), sqlQueryResultsDialog, SLOT(exec()) );
    QObject::connect( sqlQueryResultsDialog, SIGNAL(accepted()), analysisInputDialog, SLOT(exec()) );
    QObject::connect( analysisInputDialog, SIGNAL(accepted()), propagationResultsDialog, SLOT(exec()) );
    //QObject::connect( analysisInputDialog, SIGNAL(accepted()), , SLOT() );

    QObject::connect( sqlQueryDialog, SIGNAL(rejected()), &application, SLOT(quit()) );
    QObject::connect( sqlQueryResultsDialog, SIGNAL(rejected()), sqlQueryDialog, SLOT(exec()) );
    QObject::connect( analysisInputDialog, SIGNAL(rejected()), sqlQueryResultsDialog, SLOT(exec()) );
    QObject::connect( propagationResultsDialog, SIGNAL(rejected()), analysisInputDialog, SLOT(exec()) );

    sqlQueryDialog->exec();

    //result = sqlQueryDialog->exec();
    //if ( result == QDialog::Rejected )
    //{
    //    return 1;
    //}

    //result = sqlQueryResultsDialog.exec();
    //if ( result == QDialog::Rejected )
    //{
    //    return 2;
    //}

    //result = analysisInputDialog.exec();
    //if ( result == QDialog::Rejected )
    //{
    //    return 3;
    //}

    return application.exec();
}
