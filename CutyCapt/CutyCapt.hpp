#include <QMainWindow>
#include <QtWebKit>

class MainWindow : public QMainWindow {
  Q_OBJECT

public:
  MainWindow();
  void setOutputFilePath(const char* path);
  void setRenderingDelay(int ms);

private slots:
  void DocumentComplete();
  void Timeout();
  void Delayed();

private:
  void saveSnapshot();

protected:
  const char* mOutput;
  int mDelay;

};

class CutyPage : public QWebPage {
  Q_OBJECT

  
};
