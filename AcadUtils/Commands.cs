/* Vadim Semenov (c) 2020
 * Email: 5587394@mail.ru
 * Email: vad.s.semenov@gmail.com
 * Skype: semenov.vadim
 */
using System;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsInterface;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;

namespace AcadUtils
{
    partial class Main
    {

    /// <summary>
    /// Меняет все шрифты на один шрифт во всех текстовых стилях.
    /// </summary>
    [CommandMethod("AcadUtils_TextStyleChangeText", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public static void TextStyleChangeText()
        {
            //https://adn-cis.org/forum/index.php?topic=7669.0
            //Database database = HostApplicationServices.WorkingDatabase;
            Autodesk.AutoCAD.ApplicationServices.Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = acDoc.Editor;// Application.DocumentManager.MdiActiveDocument.Editor;
            Database database = acDoc.Database;
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                editor.WriteMessage("\nЗамена всех шрифтов на один во всех текстовых стилях.\n");

                PromptStringOptions promptStringOptions = new PromptStringOptions("Введите название шрифта(например: ISOCPEUR ):");
                PromptResult promptResult = editor.GetString(promptStringOptions);
                string fontName = "ISOCPEUR";

                if (!promptResult.StringResult.Equals(""))
                {
                   fontName = promptResult.StringResult;
                }

                SymbolTable symTable = (SymbolTable)transaction.GetObject(database.TextStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in symTable)
                {
                    TextStyleTableRecord symbol = (TextStyleTableRecord)transaction.GetObject(id, OpenMode.ForWrite);

                    try
                    {
                        //TODO: Access to the symbol
                        symbol.Font = new FontDescriptor(fontName, false, false, 0, 0); //symbol.FileName = "ISOCPEUR.ttf";
                        symbol.ObliquingAngle = 0;
                        editor.WriteMessage(string.Format("\nName: {0} - {1}", symbol.Name, symbol.FileName));
                    }
                    catch(Autodesk.AutoCAD.Runtime.Exception e) {
                        editor.WriteMessage("Ошибка:{0}", e.Message);
                    }

                }

                //Screen refresh
                //acDoc.TransactionManager.EnableGraphicsFlush(true);
                //acDoc.TransactionManager.QueueForGraphicsFlush();
                //Autodesk.AutoCAD.Internal.Utils.FlushGraphics();

                transaction.Commit();
            }
        }


        /// <summary>
        /// Вычисляет сумму чисел из текстовых примитивов и вставляет результат в текстовый примитив.
        /// </summary>
        [CommandMethod("AcadUtils_TextSumm", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public static void TextSumm()
        {
            //https://adn-cis.org/forum/index.php?topic=6085.0
            double text, summa = 0;

            // Получение текущего документа и базы данных
            Autodesk.AutoCAD.ApplicationServices.Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                editor.WriteMessage(Environment.NewLine + "Вычисление суммы чисел из текстовых примитивов и вставка результата в Мтекст или Текст." + Environment.NewLine);

                // Выбор объектов
                PromptSelectionOptions selectionOptions = new PromptSelectionOptions();
                selectionOptions.Keywords.Add("Выберите текстовые значения для сложения:");
				PromptSelectionResult promtResult = editor.GetSelection(selectionOptions);
                if (promtResult.Status == PromptStatus.OK)
                {
                    SelectionSet selSet = promtResult.Value;

                    RXClass dbTextRx = RXClass.GetClass(typeof(DBText));
                    RXClass mTextRx = RXClass.GetClass(typeof(MText));

                    //Прохождение по объектам выборки
                    foreach (SelectedObject acSelObj in selSet)
                    {
                        DBObject obj = acTrans.GetObject(acSelObj.ObjectId, OpenMode.ForRead);

                        //Сложение для DBText
                        if (obj.GetRXClass().Equals(dbTextRx))
                        {
                            DBText dbtext = acTrans.GetObject(obj.ObjectId, OpenMode.ForRead) as DBText;
                            if (Double.TryParse(dbtext.TextString.Replace(',', '.'), out text))
                            {
                                summa += text;
                            }
                            else { editor.WriteMessage("Число неправильного формата." + Environment.NewLine); }
                        }

                        //Сложение для MText
                        if (obj.GetRXClass().Equals(mTextRx))
                        {
                            MText mtext = acTrans.GetObject(obj.ObjectId, OpenMode.ForRead) as MText;
                            if (Double.TryParse(mtext.Text.Replace(',', '.'), out text))
                            {
                                summa += text;
                            }
                            else { editor.WriteMessage("Число неправильного формата." + Environment.NewLine); }
                        }
                    }
                    editor.WriteMessage("Сумма равна:" + summa.ToString() + Environment.NewLine);

                    //------------------------------Вставка числа-----------------------------------

                    // Выбор объектов 
                    PromptEntityResult prResult = editor.GetEntity("Выберите объект для вставки суммы:");

                    if (prResult.Status == PromptStatus.OK)
                    {
                        DBObject dbobjText = acTrans.GetObject(prResult.ObjectId, OpenMode.ForWrite);
                        //Вставка в DbText
                        if (dbobjText.GetRXClass().Equals(dbTextRx))
                        {
                            DBText dbtext = acTrans.GetObject(dbobjText.ObjectId, OpenMode.ForWrite) as DBText;
                            dbtext.TextString = summa.ToString();
                        }
                        //Вставка в MText
                        if (dbobjText.GetRXClass().Equals(mTextRx))
                        {
                            MText mtext = acTrans.GetObject(dbobjText.ObjectId, OpenMode.ForWrite) as MText;
                            mtext.Contents = summa.ToString();
                        }
                    }
                    acTrans.Commit();
                }
            }
        }


        /// <summary>
        /// Вычисляет произведение чисел из текстовых примитивов и вставляет результат в текстовый примитив.
        /// </summary>
        [CommandMethod("AcadUtils_TextMult", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public static void TextMult()
        {
            double text, mult = 1;

            // Получение текущего документа и базы данных
            Autodesk.AutoCAD.ApplicationServices.Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {


                editor.WriteMessage(Environment.NewLine + "Вычисление произведение чисел из текстовых примитивов и вставка результата в Мтекст или Текст."+ Environment.NewLine);

                // Выбор объектов 
                PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
                promptSelectionOptions.Keywords.Add("Выберите текстовые значения для умножения:");
                PromptSelectionResult promtResult = editor.GetSelection(promptSelectionOptions);
                // editor.WriteMessage(Environment.NewLine + "Выберите значения для сложения:");
                if (promtResult.Status == PromptStatus.OK)
                {
                    SelectionSet selSet = promtResult.Value;

                    RXClass dbTextRx = RXClass.GetClass(typeof(DBText));
                    RXClass mTextRx = RXClass.GetClass(typeof(MText));

                    //Прохождение по объектам выборки
                    foreach (SelectedObject acSelObj in selSet)
                    {
                        DBObject obj = acTrans.GetObject(acSelObj.ObjectId, OpenMode.ForRead);

                        //Сложение для DBText
                        if (obj.GetRXClass().Equals(dbTextRx))
                        {
                            DBText dbtext = acTrans.GetObject(obj.ObjectId, OpenMode.ForRead) as DBText;
                            if (Double.TryParse(dbtext.TextString.Replace(',', '.'), out text))
                            {
                                if (text == 0) {
                                    mult = 0;
                                    break;
                                }
                                editor.WriteMessage("Добавлен объект." + Environment.NewLine);
                                mult *= text;
                            }
                            else { editor.WriteMessage("Число неправильного формата." + Environment.NewLine); }
                        }

                        //Умножение для MText
                        if (obj.GetRXClass().Equals(mTextRx))
                        {
                            MText mtext = acTrans.GetObject(obj.ObjectId, OpenMode.ForRead) as MText;
                            if (Double.TryParse(mtext.Text.Replace(',', '.'), out text))
                            {
                                if (text == 0)
                                {
                                    mult = 0;
                                    break;
                                }
                                editor.WriteMessage("Добавлен объект." + Environment.NewLine);
                                mult *= text; 
                            }
                            else { editor.WriteMessage("Число неправильного формата." + Environment.NewLine); }
                        }
                    }
                    editor.WriteMessage("Произведение равно:" + mult.ToString() + Environment.NewLine);
                    //------------------------------Вставка числа-----------------------------------
                    // Выбор объектов 
                    PromptEntityResult prResult = editor.GetEntity("Выберите объект для вставки результата умножения:");
                    if (prResult.Status == PromptStatus.OK)
                    {
                        DBObject dbobjText = acTrans.GetObject(prResult.ObjectId, OpenMode.ForWrite);
                        //Вставка в DbText
                        if (dbobjText.GetRXClass().Equals(dbTextRx))
                        {
                            DBText dbtext = acTrans.GetObject(dbobjText.ObjectId, OpenMode.ForWrite) as DBText;
                            dbtext.TextString = mult.ToString();
                        }
                        //Вставка в MText
                        if (dbobjText.GetRXClass().Equals(mTextRx))
                        {
                            MText mtext = acTrans.GetObject(dbobjText.ObjectId, OpenMode.ForWrite) as MText;
                            mtext.Contents = mult.ToString();
                        }
                    }
                    acTrans.Commit();
                }
            }
        }


        /// <summary>
        /// Удаляет форматирование МТекста
        /// </summary>
        [CommandMethod("AcadUtils_MTextClear", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public static void ClearMtextFormatting()
        {
            Autodesk.AutoCAD.ApplicationServices.Document doc = Application.DocumentManager.MdiActiveDocument;
            Database workingDatabase = doc.Database;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionResult promptSelectionResult = editor.SelectImplied();
            if (promptSelectionResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage(Environment.NewLine + "Сброс форматирования Мтекста." + Environment.NewLine);

                PromptSelectionOptions options = new PromptSelectionOptions();
                options.MessageForAdding = "Выберите многострочный текст:";
                SelectionFilter filter = new SelectionFilter(new TypedValue[1]
                {
          new TypedValue((int)DxfCode.Start, "MTEXT")
                });
                promptSelectionResult = editor.GetSelection(options, filter);
            }
            if (promptSelectionResult.Status != PromptStatus.OK)
                return;
            Transaction transaction = workingDatabase.TransactionManager.StartTransaction();
            try
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId, OpenMode.ForRead);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(workingDatabase.CurrentSpaceId, OpenMode.ForWrite);
                foreach (ObjectId objectId in promptSelectionResult.Value.GetObjectIds())
                {
                    MText mtext = workingDatabase.TransactionManager.GetObject(objectId, OpenMode.ForWrite, false, true) as MText;
                    if (mtext != null)
                        mtext.Contents = mtext.Text;
                }
                transaction.Commit();
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("Ошибка:{0} ", ex.Message);
            }
            finally
            {
                transaction.Dispose();
            }
        }
    }
}
