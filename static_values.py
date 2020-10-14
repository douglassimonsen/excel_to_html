COLOR_INDEX = (
    '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', #0-4
    '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF', #5-9
    '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', #10-14
    '0000FFFF', '00800000', '00008000', '00000080', '00808000', #15-19
    '00800080', '00008080', '00C0C0C0', '00808080', '009999FF', #20-24
    '00993366', '00FFFFCC', '00CCFFFF', '00660066', '00FF8080', #25-29
    '000066CC', '00CCCCFF', '00000080', '00FF00FF', '00FFFF00', #30-34
    '0000FFFF', '00800080', '00800000', '00008080', '000000FF', #35-39
    '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', #40-44
    '00FF99CC', '00CC99FF', '00FFCC99', '003366FF', '0033CCCC', #45-49
    '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699', #50-54
    '00969696', '00003366', '00339966', '00003300', '00333300', #55-59
    '00993300', '00993366', '00333399', '00333333',  #60-63
)
#  indices 64 and 65 are reserved for the system foreground and background colours respectively
border_style_to_width = {
    'dashDot': '1px',
    'dashed': '1px',
    'dashDotDot': '1px',
    'dotted': '1px',
    'double': '2px',
    'hair': '1px',
    'medium': '2px',
    'mediumDashDot': '2px',
    'mediumDashDotDot': '2px',
    'mediumDashed': '2px',
    'slantDashDot': '2px',
    'thick': '3px',
    'thin': '1px',
}
border_style_to_style = {
    'dashDot': 'dashed',
    'dashed': 'dotted',
    'dashDotDot': 'dashed',
    'dotted': 'dotted',
    'double': 'double',
    'hair': 'dotted',
    'medium': 'solid',
    'mediumDashDot': 'dashed',
    'mediumDashDotDot': 'dashed',
    'mediumDashed': 'dashed',
    'slantDashDot': 'dashed',
    'thick': 'solid',
    'thin': 'solid',
}
