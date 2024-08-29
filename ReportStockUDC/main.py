from .ExtractStockUDC import ExtractFile
from .FormatStockUDC import (
    FormatBox,
    FormatI2,
    FormatVirtualLoc,
    FormatMissions,
    FormatStorage,
    FormatQuality,
    FormatRTL,
    FormatSHTLoad,
    FormatTRS,
    FormatContainer,
    FormatRej,
)


def main():

    ExtractFile()
    FormatBox()
    FormatI2()
    FormatVirtualLoc()
    FormatMissions()
    FormatStorage()
    FormatQuality()
    FormatRTL()
    FormatSHTLoad()
    FormatTRS()
    FormatContainer()
    FormatRej()


if __name__ == "__main__":
    main()
