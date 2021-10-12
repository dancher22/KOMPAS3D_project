using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KAPITypes;
using Kompas6API5;
using Kompas6Constants;

namespace Основная_надпись
{
    public class Работа_с_Компасом
    {
        public void Фамилия_В_Штамп(KompasObject kompas, ksDocument2D doc2D, int index, string surname)
        {
            ksStamp stamp = (ksStamp)doc2D.GetStamp();
            if (stamp != null)
            {
                if (stamp.ksOpenStamp() == 1)
                {
                    switch (index)
                    {
                        case 0:
                            stamp.ksColumnNumber(110);
                            break;
                        case 1:
                            stamp.ksColumnNumber(111);
                            break;
                        case 2:
                            stamp.ksColumnNumber(112);
                            break;
                        case 3:
                            stamp.ksColumnNumber(113);
                            break;
                        case 4:
                            stamp.ksColumnNumber(114);
                            break;
                        case 5:
                            stamp.ksColumnNumber(115);
                            break;
                        case 6:
                            stamp.ksColumnNumber(10);
                            break;
                    }
                    ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
                    if (itemParam != null)
                    {
                        itemParam.Init();
                        ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                        if (itemFont != null)
                        {
                            itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                            itemParam.s = surname;
                            doc2D.ksTextLine(itemParam);
                        }
                    }
                }
                stamp.ksCloseStamp();
            }
        }

        public void Дата_В_Штамп(KompasObject kompas, ksDocument2D doc2D, int index, string curDate)
        {
            ksStamp stamp = (ksStamp)doc2D.GetStamp();
            if (stamp != null)
            {
                if (stamp.ksOpenStamp() == 1)
                {
                    switch (index)
                    {
                        case 0:
                            stamp.ksColumnNumber(130);
                            break;
                        case 1:
                            stamp.ksColumnNumber(131);
                            break;
                        case 2:
                            stamp.ksColumnNumber(132);
                            break;
                        case 3:
                            stamp.ksColumnNumber(133);
                            break;
                        case 4:
                            stamp.ksColumnNumber(134);
                            break;
                        case 5:
                            stamp.ksColumnNumber(135);
                            break;
                    }
                    ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
                    if (itemParam != null)
                    {
                        itemParam.Init();
                        ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                        if (itemFont != null)
                        {
                            itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                            itemParam.s = curDate;
                            doc2D.ksTextLine(itemParam);
                        }
                    }
                }
                stamp.ksCloseStamp();
            }
        }

        public void ВставкаПодписи(KompasObject kompas, ksDocument2D doc, string путьКподписи, int width, int index)
        {
            ksRasterParam raster = (ksRasterParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RasterParam);
            raster.Init();
            raster.embeded = false;
            raster.fileName = путьКподписи;

            if (путьКподписи != null)
            {
                double scale = doc.ksCalcRasterScale(raster.fileName, 15, 5);

                ksPlacementParam param = (ksPlacementParam)kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam);
                param.Init();
                param.angle = 0;
                param.scale_ = scale;
                param.xBase = width - 149;
                switch (index)
                {
                    case 0:
                        param.yBase = 30;
                        break;
                    case 1:
                        param.yBase = 25;
                        break;
                    case 2:
                        param.yBase = 20;
                        break;
                    case 3:
                        param.yBase = 15;
                        break;
                    case 4:
                        param.yBase = 10;
                        break;
                    case 5:
                        param.yBase = 5;
                        break;
                }
                raster.SetPlace(param);
                doc.ksInsertRaster(raster);
            }
        }

        public int ReturnWidth(KompasObject kompas, ksDocument2D doc2D1)
        {
            ksDocumentParam DocumentParam = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
            doc2D1.ksGetObjParam(doc2D1.reference, DocumentParam, ldefin2d.ALLPARAM);
            ksSheetPar sheet = (ksSheetPar)DocumentParam.GetLayoutParam();
            int xFormat = 0;
            if (sheet != null)
            {
                ksStandartSheet iStandartSheet = (ksStandartSheet)sheet.GetSheetParam();
                if (iStandartSheet != null)
                {
                    if (iStandartSheet.direct == true)
                    {
                        int[,] intArray = new int[5, 2] { { 1189, 841 }, { 841, 594 }, { 594, 420 }, { 420, 297 }, { 297, 210 } };
                        if (iStandartSheet.multiply == 1)
                            xFormat = intArray[(int)iStandartSheet.format, 0];
                        else
                            xFormat = intArray[(int)iStandartSheet.format, 0] * (iStandartSheet.multiply) * 210 / 297;
                    }
                    else
                    {
                        int[,] intArray = new int[5, 2] { { 841, 1189 }, { 594, 841 }, { 420, 594 }, { 297, 420 }, { 210, 297 } };
                        if (iStandartSheet.multiply == 1)
                            xFormat = intArray[(int)iStandartSheet.format, 0];
                        else
                            xFormat = intArray[(int)iStandartSheet.format, 0] * (iStandartSheet.multiply) * 297 / 210;
                    }
                }
            }
            return xFormat;
        }
    }
}
