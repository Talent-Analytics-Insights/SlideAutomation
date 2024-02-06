def first_row(display: bool = True, text_color: str = "FFFFFF"):
    """
    Get first row style element with text color. Default is white (FFFFFF).
    """
    if display:
        # Return style XML for first row table with the correct text color
        return f"""
            <a:firstRow>
            <a:tcTxStyle b="on">
                <a:fontRef idx="minor">
                    <a:prstClr val="black" />
                </a:fontRef>
                <a:srgbClr val="{text_color}" />
            </a:tcTxStyle>
            <a:tcStyle>
                <a:tcBdr>
                    <a:bottom>
                        <a:ln w="38100" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:bottom>
                </a:tcBdr>
            </a:tcStyle>
        </a:firstRow>
        """
    else:
        return ""


def get_style_element(style_id: str, style_name: str, bg_color: str, even_color: str, odd_color: str, include_first_row: bool = True):
    """
    Return Table style element with correct colors for background, even rows, and odd rows, possibly including top row with different value
    """
    return """
    <a:tblStyle styleId="{}" styleName="{}">
        <a:wholeTbl>
            <a:tcTxStyle>
                <a:fontRef idx="minor">
                    <a:prstClr val="black" />
                </a:fontRef>
                <a:schemeClr val="dk1" />
            </a:tcTxStyle>
            <a:tcStyle>
                <a:tcBdr>
                    <a:left>
                        <a:ln w="12700" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:left>
                    <a:right>
                        <a:ln w="12700" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:right>
                    <a:top>
                        <a:ln w="12700" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:top>
                    <a:bottom>
                        <a:ln w="12700" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:bottom>
                    <a:insideH>
                        <a:ln w="12700" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:insideH>
                    <a:insideV>
                        <a:ln w="12700" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:insideV>
                </a:tcBdr>
                <a:fill>
                    <a:solidFill>
                        <a:srgbClr val="{}" />
                    </a:solidFill>
                </a:fill>
            </a:tcStyle>
        </a:wholeTbl>
        <a:band1H>
            <a:tcStyle>
                <a:tcBdr />
                <a:fill>
                    <a:solidFill>
                        <a:srgbClr val="{}" />
                    </a:solidFill>
                </a:fill>
            </a:tcStyle>
        </a:band1H>
        <a:band2H>
            <a:tcStyle>
                <a:tcBdr />
                <a:fill>
                    <a:solidFill>
                        <a:srgbClr val="{}" />
                    </a:solidFill>
                </a:fill>
            </a:tcStyle>
        </a:band2H>
        <a:band1V>
            <a:tcStyle>
                <a:tcBdr />
            </a:tcStyle>
        </a:band1V>
        <a:band2V>
            <a:tcStyle>
                <a:tcBdr />
            </a:tcStyle>
        </a:band2V>

        <a:lastRow>
            <a:tcTxStyle b="on">
                <a:fontRef idx="minor">
                    <a:prstClr val="black" />
                </a:fontRef>
                <a:schemeClr val="lt1" />
            </a:tcTxStyle>
            <a:tcStyle>
                <a:tcBdr>
                    <a:top>
                        <a:ln w="38100" cmpd="sng">
                            <a:solidFill>
                                <a:schemeClr val="lt1" />
                            </a:solidFill>
                        </a:ln>
                    </a:top>
                </a:tcBdr>
            </a:tcStyle>
        </a:lastRow>
    {}
    </a:tblStyle>
    """.format(style_id, style_name, bg_color, even_color, odd_color, first_row(display=include_first_row))
