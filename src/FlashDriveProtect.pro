TEMPLATE = app

CONFIG(release, debug|release):DEFINES += NDEBUG

QMAKE_LFLAGS += -static
QMAKE_CXXFLAGS += -Wpedantic

LIBS += -lgdi32 -lcomctl32

SOURCES += main.cpp

RC_FILE = res.rc
