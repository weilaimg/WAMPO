
function Get_HighlightColor(wdColorIndex){

    var RGB_Color = 0
    var R = 0;
    var G = 0;
    var B = 0;
    
    switch(wdColorIndex)
    {

        case wps.Enum.wdNoHighlight:
            {
                R = 255;
                G = 255;
                B = 255;
                break;
            }
        case wps.Enum.wdBlack:
            {
                R = 0;
                G = 0;
                B = 0;
                break;
            }
        case wps.Enum.wdBlue:
            {
                R = 0;
                G = 0;
                B = 255;
                break;
            }
        case wps.Enum.wdBrightGreen:
            {
                R = 0;
                G = 255;
                B = 0;
                break;
            }
        case wps.Enum.wdGray25:
            {
                R = 192;
                G = 192;
                B = 192;
                break;
            }
        case wps.Enum.wdGray50:
            {
                R = 128;
                G = 128;
                B = 128;
                break;
            }
        case wps.Enum.wdGreen:
            {
                R = 0;
                G = 128;
                B = 0;
                break;
            }
        case wps.Enum.wdPink:
            {
                R = 255;
                G = 0;
                B = 255;
                break;
            }
        case wps.Enum.wdYellow:
            {
                R = 255;
                G = 255;
                B = 0;
                break;
            }
        case wps.Enum.wdDarkBlue:
            {
                R = 0;
                G = 0;
                B = 128;
                break;
            }
        case wps.Enum.wdDarkRed:
            {
                R = 128;
                G = 0;
                B = 0;
                break;
            }
        case wps.Enum.wdDarkYellow:
            {
                R = 128;
                G = 128;
                B = 0;
                break;
            }
        case wps.Enum.wdRed:
            {
                R = 255;
                G = 0;
                B = 0;
                break;
            }
        case wps.Enum.wdTeal:
            {
                R = 0;
                G = 128;
                B = 128;
                break;
            }
        case wps.Enum.wdTurquoise:
            {
                R = 0;
                G = 255;
                B = 255;
                break;
            }
        case wps.Enum.wdViolet:
            {
                R = 128;
                G = 0;
                B = 128;
                break;
            }
        case wps.Enum.wdWhite:
            {
                R = 255;
                G = 255;
                B = 255;
                break;
            }
        default:
            {
                return -1
            }

    }

    RGB_Color = (R<<16) | (G<<8) | B

    return RGB_Color


}
