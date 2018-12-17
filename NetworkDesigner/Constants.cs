using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetworkDesigner
{

    public static class Constants
    {
        public const int LV_EXCEL_ARRAY_SIZE = 20;
        public const double FENCE_DIAMETER = 3;
        public const double MSTN_TO_RMD_COORD_MAPPING_SCALE = 4;
        public const string TRFRZONE_SHAPE_LEVELNAME = "Level 11";
        public const string AIRDAC_LEVEL_NAME = "Level 18";
        public const string LV_LEVEL_NAME = "Level 25";
        public const string MV_LEVEL_NAME = "Level 29";
        public const string AIRDAC_NAME = "Airdac";
        public const string ABC35_RED = "ABC35-R";
        public const string ABC35_WHITE = "ABC35-W";
        public const string ABC35_BLUE = "ABC35-B";
        public const string ABC35_REDWHITE = "ABC35-RW";
        public const string ABC35_WHITEBLUE = "ABC35-WB";
        public const string ABC35_BLUERED = "ABC35-RB";

        public const string ARVIDAL_RED = "35-R";
        public const string ARVIDAL_WHITE = "35-W";
        public const string ARVIDAL_BLUE = "35-B";
        public const string ARVIDAL_REDWHITE = "35-RW";
        public const string ARVIDAL_WHITEBLUE = "35-WB";
        public const string ARVIDAL_BLUERED = "35-RB";

        public const string OAK_LINESTYLE = "OAK";
        public const string ARVIDAL_LINESTYLE = "35";
        public const string AIRDAC_LINESTYLE = "4mm AD";
        public const string AIRDAC_A_LSTYLE = "4mm AD A";
        public const string AIRDAC_B_LSTYLE = "4mm AD B";
        public const string AIRDAC_C_LSTYLE = "4mm AD C";

        // text levels
        public const string CBAZZ_LEVELNAME = "Level 55";

        public const string MVTEXT_LEVELNAME = "Level 56";
        public const string LVTEXT_LEVELNAME = "Level 57";
        public const string TRFRTEXT_LEVELNAME = "Level 58";

        // auxilliary types levels
        public const string TRFR_CELLNAME = "TXTCS";

        public const string LVPOLES_CELLNAME = "009";
        public const string MVPOLES_CELLNAME = "010";
        public const string MVLVPOLES_CELLNAME = "011";
        public const string MVSTAY_CELLNAME = "003";
        public const string LVSTAY_CELLNAME = "001";
        public const string LVSTRUT_CELLNAME = "002";
        public const string MVSTRUT_CELLNAME = "004";

        public const string MVSTAY_LEVELNAME = "Level 35";
        public const string LVSTAY_LEVELNAME = "Level 34";
        public const string MV_FLYSTAY_CELLNAME = "017";
        public const string LV_FLYSTAY_CELLNAME = "015";
        public const string FENCE_SELECTION_NEW = "FENCE SELECTION NEW";
        public const string FENCE_SELECTION_APPEND = "FENCE SELECTION APPEND";
        public const string STRUCTURES_TEXT_TYPE = "TEXT";

        public const bool SINGLE_PHASE_LV = true;

        public const int XCELL_OFFSET = 5;
        public static bool selectbySettingsDone = false;

        public static string[] LV_1PH_STYLE_LIST = new[] { ABC35_BLUE, ABC35_RED, ABC35_WHITE };
        public static string[] LV_2PH_STYLE_LIST = new[] { ABC35_BLUERED, ABC35_REDWHITE, ABC35_WHITEBLUE };

        public static string[] LSTYLE_LIST = new[] { "g", "j" };

        public static string[] LV_1PHStyles = new[] { ABC35_BLUE, ABC35_RED, ABC35_WHITE };

        public static string[] LV_2PHStyles = new[] { ABC35_WHITEBLUE, ABC35_BLUERED, ABC35_REDWHITE };
        public static string[] Airdac_Styles = new[] { AIRDAC_A_LSTYLE, AIRDAC_B_LSTYLE, AIRDAC_C_LSTYLE };
    }
}