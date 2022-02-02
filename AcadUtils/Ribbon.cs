/* Vadim Semenov (c) 2020
 * Email: 5587394@mail.ru
 * Email: vad.s.semenov@gmail.com
 * Skype: semenov.vadim
 */
using System;
using System.Windows.Media.Imaging;
using System.Reflection;
using Autodesk.Windows;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcadUtils
{
    partial class Main
    {
        BitmapImage getBitmap(string fileName)
        {
            BitmapImage bmp = new BitmapImage();
            // BitmapImage.UriSource must be in a BeginInit/EndInit block.              
            bmp.BeginInit();
            bmp.UriSource = new Uri(string.Format("pack://application:,,,/{0};component/{1}",
              Assembly.GetExecutingAssembly().GetName().Name,
              fileName));
            bmp.EndInit();

            return bmp;
        }


        
        /// <summary>
        /// Построение вкладки
        /// </summary>
        [CommandMethod("AcadUtils_CreateRibbon")]
       public void BuildRibbonTab()
        {
            // Если лента еще не загружена
            if (!isLoaded())
            {
                // Строим вкладку
                AddRibbon();
                // Подключаем обработчик событий изменения системных переменных
                acApp.SystemVariableChanged += new SystemVariableChangedEventHandler(acadApp_SystemVariableChanged);
            }
        }
       
        
        /// <summary>
        /// Проверка "загруженности" ленты
        /// </summary>
        /// <returns></returns>
         bool isLoaded()
        {
            bool _loaded = false;
            RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
            // Делаем итерацию по вкладкам ленты
            foreach (RibbonTab tab in ribCntrl.Tabs)
            {
                // И если у вкладки совпадает идентификатор и заголовок, то значит вкладка загружена
                if (tab.Id.Equals("ACADUTILS_RIBBON_TAB_ID") & tab.Title.Equals("Утилиты Autocad"))
                { _loaded = true; break; }
                else _loaded = false;
            }
            return _loaded;
        }



         /// <summary>
         /// Удаление своей вкладки с ленты
         /// </summary>
        [CommandMethod("AcadUtils_RemoveRibbon")]
       public void RemoveRibbonTab()
        {
            try
            {
                RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
                // Делаем итерацию по вкладкам ленты
                foreach (RibbonTab tab in ribCntrl.Tabs)
                {
                    if (tab.Id.Equals("ACADUTILS_RIBBON_TAB_ID") & tab.Title.Equals("Утилиты Autocad"))
                    {
                        // И если у вкладки совпадает идентификатор и заголовок, то удаляем эту вкладку
                        ribCntrl.Tabs.Remove(tab);
                        // Отключаем обработчик событий
                        acApp.SystemVariableChanged -= new SystemVariableChangedEventHandler(acadApp_SystemVariableChanged);
                        // Останавливаем итерацию
                        break;
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }



         /// <summary>
         /// Обработка события изменения системной переменной. 
         /// Будем следить за системной переменной WSCURRENT (текущее рабочее пространство),
         /// чтобы наша вкладка не "терялась" при изменение рабочего пространства
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
        void acadApp_SystemVariableChanged(object sender, SystemVariableChangedEventArgs e)
        {
            if (e.Name.Equals("WSCURRENT")) BuildRibbonTab();
        }

        
        public void AddRibbon()
        {
            Autodesk.Windows.RibbonControl ribbonControl = Autodesk.Windows.ComponentManager.Ribbon;

            // Создаем Ribbon tab
            RibbonTab Tab = new RibbonTab();
            Tab.Title = "Утилиты Autocad";
            Tab.Id = "ACADUTILS_RIBBON_TAB_ID";

            ribbonControl.Tabs.Add(Tab);

            // Создаем Ribbon panel
            Autodesk.Windows.RibbonPanelSource srcPanel = new RibbonPanelSource();
            srcPanel.Title = "Форматирование текста";

            RibbonPanel Panel = new RibbonPanel();
            Panel.Source = srcPanel;
            Tab.Panels.Add(Panel);

            // Создаем кнопки
            Autodesk.Windows.RibbonButton button1 = new RibbonButton();
            button1.Text = "Замена текстовых стилей";
            button1.Id = "cmdButton1";
            button1.Size = RibbonItemSize.Large;
            button1.ShowText = true;
            button1.ShowImage = true;
            button1.LargeImage = getBitmap("TextStReset_large.png");
            button1.Image = getBitmap("TextStReset_small.png");
            button1.CommandParameter = "AcadUtils_TextStyleChangeText";
            button1.CommandHandler = new MyCmdHandler();

            Autodesk.Windows.RibbonButton button2 = new RibbonButton();
            button2.Text = "Сброс форматирования Мтекст";
            button2.Id = "cmdButton2";
            button2.Size = RibbonItemSize.Large;
            button2.ShowText = true;
            button2.ShowImage = true;
            button2.LargeImage = getBitmap("MTextReset_large.png");
            button2.Image = getBitmap("MTextReset_small.png");
            button2.CommandParameter = "AcadUtils_MTextClear";
            button2.CommandHandler = new MyCmdHandler();

            Autodesk.Windows.RibbonButton button3 = new RibbonButton();
            button3.Text = "Сложение";
            button3.Id = "cmdButton3";
            button3.Size = RibbonItemSize.Standard;
            button3.ShowText = true;
            button3.ShowImage = true;
            button3.LargeImage = getBitmap("TextSumm_large.png");
            button3.Image = getBitmap("TextSumm_small.png");
            button3.CommandParameter = "AcadUtils_TextSumm";
            button3.CommandHandler = new MyCmdHandler();

            Autodesk.Windows.RibbonButton button4 = new RibbonButton();
            button4.Text = "Умножение";
            button4.Id = "cmdButton4";
            button4.Size = RibbonItemSize.Standard;
            button4.ShowText = true;
            button4.ShowImage = true;
            button4.LargeImage = getBitmap("TextMult_large.png");
            button4.Image = getBitmap("TextMult_small.png");
            button4.CommandParameter = "AcadUtils_TextMult";
            button4.CommandHandler = new MyCmdHandler();

            // Создаем Split button
            RibbonSplitButton ribSplitButton = new RibbonSplitButton();
            ribSplitButton.Text = "RibbonSplitButton"; //Required not to crash AutoCAD when using cmd locator
            ribSplitButton.ShowText = true;

            ribSplitButton.Items.Add(button1);
            ribSplitButton.Items.Add(button2);

            ////Создаем RowPanel
            RibbonRowPanel ribRowPanel = new RibbonRowPanel();
            ribRowPanel.Items.Add(button3);
            ribRowPanel.Items.Add(button4);
            
            RibbonRowBreak ribbonRowBreak = new RibbonRowBreak();
            ribRowPanel.Items.Add(ribbonRowBreak);

            ribRowPanel.Items.Add(ribSplitButton);


            //srcPanel.Items.Add(ribSplitButton);
            //srcPanel.Items.Add(new RibbonSeparator());
            //srcPanel.Items.Add(button3);
            //srcPanel.Items.Add(button4);
            srcPanel.Items.Add(ribRowPanel);


            //Tab.IsActive = true;
        }

        public class MyCmdHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged;

            public void Execute(object parameter)
            {
                Document doc = acApp.DocumentManager.MdiActiveDocument;

                if (parameter is RibbonButton)
                {
                    RibbonButton button = parameter as RibbonButton;

                    if (button.Id != null) //&& button.Id.Equals("cmdButton1")
                    {
                        //Кроме того, на AutoCAD 2015(и новее) вы можете использовать Editor.Command или Editor.CommandAsync, что намного лучше.
                        doc.SendStringToExecute(button.CommandParameter + " ", true, false, true);
                    }

                }
            }
        }

    }
}
