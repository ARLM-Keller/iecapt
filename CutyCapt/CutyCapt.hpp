#include <QMainWindow>
#include <QtWebKit>

class MainWindow : public QMainWindow {
  Q_OBJECT

public:
  MainWindow();
  void setOutputFilePath(char* path);
  void setRenderingDelay(int ms);

private slots:
  void DocumentComplete();
  void Timeout();
  void Delayed();

private:
  void saveSnapshot();

protected:
  char* mOutput;
  int mDelay;

};

