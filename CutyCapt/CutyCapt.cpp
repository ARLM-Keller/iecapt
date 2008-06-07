////////////////////////////////////////////////////////////////////
//
// CutyCapt - A Qt WebKit Web Page Rendering Capture Utility
//
// Copyright (C) 2003-2008 Bjoern Hoehrmann <bjoern@hoehrmann.de>
//
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// $Id$
//
////////////////////////////////////////////////////////////////////

#include <QApplication>
#include <QMainWindow>
#include <QtWebKit>
#include <QtGui>
#include <QSvgGenerator>
#include <QPrinter>
#include <QTimer>
#include "CutyCapt.hpp"

Q_IMPORT_PLUGIN(qjpeg)
Q_IMPORT_PLUGIN(qgif)
Q_IMPORT_PLUGIN(qtiff)
Q_IMPORT_PLUGIN(qsvg)
Q_IMPORT_PLUGIN(qmng)
Q_IMPORT_PLUGIN(qico)

MainWindow::MainWindow() {

  QWebView *browser = new QWebView(this);
  setCentralWidget(browser);
  mOutput = NULL;
  mDelay = 0;

}

void
MainWindow::setOutputFilePath(char* path) {
  mOutput = path;
}

void
MainWindow::setRenderingDelay(int ms) {
  mDelay = ms;
}

void
MainWindow::DocumentComplete() {

  if (mDelay > 0) {
    QTimer::singleShot(mDelay, this, SLOT(Delayed()));
    return;
  }

  saveSnapshot();
  QApplication::exit();
}

void
MainWindow::Timeout() {
  saveSnapshot();
  QApplication::exit();
}

void
MainWindow::Delayed() {
  saveSnapshot();
  QApplication::exit();
}

void
MainWindow::saveSnapshot() {
  QWebView *view = qobject_cast<QWebView *>(centralWidget());
  QWebPage *page = view->page();
  QWebFrame *mainFrame = page->mainFrame();
  char* ext = strrchr(mOutput, '.');

  mainFrame->setScrollBarPolicy(Qt::Horizontal, Qt::ScrollBarAlwaysOff);
  mainFrame->setScrollBarPolicy(Qt::Vertical, Qt::ScrollBarAlwaysOff);

  page->setViewportSize( mainFrame->contentsSize() );

  QPainter painter;

  if (ext && strcmp(".svg", ext) == 0) {
    QSvgGenerator svg;
    svg.setFileName(mOutput);
    svg.setSize(page->viewportSize());
    painter.begin(&svg);
    mainFrame->render(&painter);
    painter.end();

  } else if (ext && (strcmp(".pdf", ext) == 0
                  || strcmp(".ps", ext) == 0)) {

    QPrinter printer;
    printer.setPageSize(QPrinter::A4);
    printer.setOutputFileName(mOutput);
    mainFrame->print(&printer);

  } else {
    QImage image(page->viewportSize(), QImage::Format_ARGB32);
    painter.begin(&image);
    mainFrame->render(&painter);
    painter.end();
    image.save(mOutput);
  }
}

void
CaptHelp(void) {
  printf("%s", ""
    " -----------------------------------------------------------------------------\n"
    " Usage: cutycapt --url=http://www.example.org/ --out=localfile.png            \n"
    " -----------------------------------------------------------------------------\n"
    "  --url=<url>                    The URL to capture (http:...|file:...|...)   \n"
    "  --out=<path>                   The target file (.png|pdf|ps|svg|jpeg|...)   \n"
    "  --min-width=<int>              Minimal width for the image (default: 800)   \n"
    "  --max-wait=<ms>                Don't wait more than (default: 90000, inf: 0)\n"
    "  --delay=<ms>                   After successful load, wait (default: 0)     \n"
    "  --user-styles=<url>            Location of user style sheet, if any         \n"
    "  --javascript=<on|off>          JavaScript execution (default: on)           \n"
    "  --java=<on|off>                Java execution (default: unknown)            \n"
    "  --plugins=<on|off>             Plugin execution (default: unknown)          \n"
    "  --private-browsing=<on|off>    Private browsing (default: unknown)          \n"
    "  --auto-load-images=<on|off>    Automatic image loading (default: on)        \n"
    "  --js-can-open-windows=<on|off> Script can open windows? (default: unknown)  \n"
    "  --js-can-access-clipboard=<on|off> Script clipboard privs (default: unknown)\n"
    " -----------------------------------------------------------------------------\n"
    " http://iecapt.sf.net - (c) 2003-2008 Bjoern Hoehrmann - <bjoern@hoehrmann.de>\n"
    "");
}

void
CaptSetWebOption(QWebView* browser, QWebSettings::WebAttribute option, char* value) {
  QWebSettings *settings = browser->settings();
  
  if (strcmp(value, "on") == 0)
    settings->setAttribute(option, true);
  else if (strcmp(value, "off") == 0)
    settings->setAttribute(option, false);
  else
    (void)0; // TODO: ...
}

int
main(int argc, char *argv[]) {
  int argHelp = 0;
  int argDelay = 0;
  int argSilent = 0;
  int argMinWidth = 800;
  int argDefHeight = 600;
  int argMaxWait = 90000;

  char* argUrl = NULL;
  char* argOut = NULL;
  char* argUserStyle = NULL;
  char* argIconDbPath = NULL;

  QApplication app(argc, argv, true);

  MainWindow *main = new MainWindow();

  app.connect(main->centralWidget(),
    SIGNAL(loadFinished(bool)),
    main,
    SLOT(DocumentComplete()));

  QWebView *browser = qobject_cast<QWebView *>(main->centralWidget());

  // Parse command line parameters
  for (int ax = 1; ax < argc; ++ax) {
    size_t nlen;

    char* s = argv[ax];
    char* value;

    // boolean options
    if (strcmp("--silent", s) == 0) {
      argSilent = 1;
      continue;

    } else if (strcmp("--help", s) == 0) {
      argHelp = 1;
      break;
    } 

    value = strchr(s, '=');

    if (value == NULL) {
      // TODO: error
      argHelp = 1;
      break;
    }

    nlen = value++ - s;

    // --name=value options
    if (strncmp("--url", s, nlen) == 0) {
      argUrl = value;

    } else if (strncmp("--min-width", s, nlen) == 0) {
      // TODO: add error checking here?
      argMinWidth = (unsigned int)atoi(value);

    } else if (strncmp("--delay", s, nlen) == 0) {
      // TODO: see above
      argDelay = (unsigned int)atoi(value);

    } else if (strncmp("--max-wait", s, nlen) == 0) {
      // TODO: see above
      argMaxWait = (unsigned int)atoi(value);

    } else if (strncmp("--out", s, nlen) == 0) {
      argOut = value;

    } else if (strncmp("--user-styles", s, nlen) == 0) {
      argUserStyle = value;

    } else if (strncmp("--icon-database-path", s, nlen) == 0) {
      argIconDbPath = value;

    } else if (strncmp("--auto-load-images", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::AutoLoadImages, value);

    } else if (strncmp("--javascript", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::JavascriptEnabled, value);

    } else if (strncmp("--java", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::JavaEnabled, value);

    } else if (strncmp("--plugins", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::PluginsEnabled, value);

    } else if (strncmp("--private-browsing", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::PrivateBrowsingEnabled, value);

    } else if (strncmp("--js-can-open-windows", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::JavascriptCanOpenWindows, value);

    } else if (strncmp("--js-can-access-clipboard", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::JavascriptCanAccessClipboard, value);

    } else if (strncmp("--developer-extras", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::DeveloperExtrasEnabled, value);

    } else if (strncmp("--links-included-in-focus-chain", s, nlen) == 0) {
      CaptSetWebOption(browser, QWebSettings::LinksIncludedInFocusChain, value);

    } else {
      // TODO: error
      argHelp = 1;
    }
  }

  if (argUrl == NULL || argOut == NULL || argHelp) {
      CaptHelp();
      return EXIT_FAILURE;
  }

  main->setOutputFilePath(argOut);

  if (argMaxWait > 0) {
    QTimer::singleShot(argMaxWait, main, SLOT(Timeout()));
  }

  if (argUserStyle != NULL)
    // TODO: does this need any syntax checking?
    browser->settings()->setUserStyleSheetUrl( QUrl(argUserStyle) );

  if (argIconDbPath != NULL)
    // TODO: does this need any syntax checking?
    browser->settings()->setIconDatabasePath(argUserStyle);

  browser->resize(argMinWidth, argDefHeight);
  browser->load( QUrl(argUrl) );

  return app.exec();
}
