#! /usr/local/bin/python3

# example file for powerpoint slides generation
from pptx import Presentation
from pptx.util import Inches

# more modules as reference below
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
# chart data related
# from pptx.chart.data import ChartData
# from pptx.enum.chart import XL_TICK_MARK
# from pptx.enum.chart import XL_DATA_LABEL_POSITION
# from pptx.enum.chart import XL_LEGEND_POSITON
# from pptx.enum.chart import XL_CHART_TYPE
# from pptx.enum.chart import XL_MARKER_STYLE

# define variables
COVER_PAGE_LAYOUT = 0
DELIVERY_PAGE_LAYOUT = 1
CONTENT_PAGE_LAYOUT = 2
SCHEDULE_LAYOUT = 2

CHEVRON_WIDTH = Inches(0.8)
CHEVRON_HEIGHT = Inches(0.3)

SCHEDULE_FILL_COLOR = RGBColor(0,0,128)
SCHEDULE_LINE_COLOR = RGBColor(200,200,200)
SCHEDULE_LINE_WIDTH = Pt(2)

def create_schedule_frame(slide,pos,start,end,title):
    ''' this fucntion add schedule shapes
    '''
    left = Inches(pos[0])
    top = Inches(pos[1])
    shapes = slide.shapes
    for x in range(start,end+1):
        print("Generating",x,"blocks...")
        shape = shapes.add_shape(MSO_SHAPE.CHEVRON,left, top, CHEVRON_WIDTH, CHEVRON_HEIGHT)                       
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = SCHEDULE_FILL_COLOR
        line = shape.line
        line.color.rgb = SCHEDULE_LINE_COLOR
        line.width = SCHEDULE_LINE_WIDTH
        shape.text = title + str(x)
        left += CHEVRON_WIDTH


if __name__ == '__main__':
    '''                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    text_frame.clear()
                shape.text = "P02"
    '''
    prs = Presentation('template.pptx')
    slide_layout = prs.slide_layouts[SCHEDULE_LAYOUT]
    slide = prs.slides.add_slide(slide_layout)
    create_schedule_frame(slide,[1,1],2,10,'P')
    prs.save('schedule.pptx')
