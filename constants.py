from enum import Enum, unique


class WdColorIndex(Enum):
    wdAuto = 0  # 自动配色。 默认值；通常为黑色。
    wdBlack = 1  # 黑色。
    wdBlue = 2  # 蓝色。
    wdBrightGreen = 4  # 鲜绿色。
    wdByAuthor = -1  # 由文档作者定义的颜色。
    wdDarkBlue = 9  # 深蓝色。
    wdDarkRed = 13  # 深红色。
    wdDarkYellow = 14  # 深黄色。
    wdGray25 = 16  # 25% 灰色底纹。
    wdGray50 = 15  # 50% 灰色底纹。
    wdGreen = 11  # 绿色。
    wdNoHighlight = 0  # 清除已应用的突出显示。
    wdPink = 5  # 粉红色。
    wdRed = 6  # 红色。
    wdTeal = 10  # 青色。
    wdTurquoise = 3  # 青绿色。
    wdViolet = 12  # 紫色。
    wdWhite = 8  # 白色。
    wdYellow = 7  # 黄色。


@unique
class WdPaperSize(Enum):
    wdPaper10x14 = 0  # 10 英寸（25.4 厘米）宽，14 英寸（35.56 厘米）长。
    wdPaper11x17 = 1  # 正规 11 英寸（27.94 厘米）宽，17 英寸（43.18 厘米）长。
    wdPaperA3 = 6  # A3 尺寸。
    wdPaperA4 = 7  # A4 尺寸。
    wdPaperA4Small = 8  # 小 A4 尺寸。
    wdPaperA5 = 9  # A5 尺寸。
    wdPaperB4 = 10  # B4 尺寸。
    wdPaperB5 = 11  # B5 尺寸。
    wdPaperCSheet = 12  # C Sheet 尺寸。
    wdPaperCustom = 41  # 自定义纸张大小。
    wdPaperDSheet = 13  # D Sheet 尺寸。
    wdPaperEnvelope10 = 25  # 正规信封，尺寸 10。
    wdPaperEnvelope11 = 26  # 信封，尺寸 11。
    wdPaperEnvelope12 = 27  # 信封，尺寸 12。
    wdPaperEnvelope14 = 28  # 信封，尺寸 14。
    wdPaperEnvelope9 = 24  # 信封，尺寸 9。
    wdPaperEnvelopeB4 = 29  # B4 信封。
    wdPaperEnvelopeB5 = 30  # B5 信封。
    wdPaperEnvelopeB6 = 31  # B6 信封。
    wdPaperEnvelopeC3 = 32  # C3 信封。
    wdPaperEnvelopeC4 = 33  # C4 信封。
    wdPaperEnvelopeC5 = 34  # C5 信封。
    wdPaperEnvelopeC6 = 35  # C6 信封。
    wdPaperEnvelopeC65 = 36  # C65 信封。
    wdPaperEnvelopeDL = 37  # DL 信封。
    wdPaperEnvelopeItaly = 38  # 意大利式信封。
    wdPaperEnvelopeMonarch = 39  # 君主式信封。
    wdPaperEnvelopePersonal = 40  # 私人信封。
    wdPaperESheet = 14  # E Sheet 尺寸。
    wdPaperExecutive = 5  # 行政型尺寸。
    wdPaperFanfoldLegalGerman = 15  # 德国正规复写簿尺寸。
    wdPaperFanfoldStdGerman = 16  # 德国标准复写簿尺寸。
    wdPaperFanfoldUS = 17  # 美国复写簿尺寸。
    wdPaperFolio = 18  # 对开尺寸。
    wdPaperLedger = 19  # 分类帐尺寸。
    wdPaperLegal = 4  # 正规尺寸。
    wdPaperLetter = 2  # 信函尺寸。
    wdPaperLetterSmall = 3  # 小型信函尺寸。
    wdPaperNote = 20  # 便笺尺寸。
    wdPaperQuarto = 21  # 四开尺寸。
    wdPaperStatement = 22  # 报表尺寸。
    wdPaperTabloid = 23  # 文摘尺寸。
