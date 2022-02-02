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
using Microsoft.Win32;
using System.Reflection;
using Autodesk.Windows;
//Необходимо добавить ссылки на библиотеки:
//accoremgd.dll
//acdbmgd.dll
//acmgd.dll
//AdWindows.dll
//из папки Автокад, или из комплекта ObjectArx2013-2021
//
//Для загрузки сборки, в Автокаде ввести комманду _NETLOAD

namespace AcadUtils
{

    public partial class Main : IExtensionApplication
    {
        //Версии Net framework в зависимости от версии Autocad:
        //https://help.autodesk.com/view/OARX/2021/ENU/?guid=GUID-A6C680F2-DE2E-418A-A182-E4884073338A


        /// <summary>
        /// Регистрирует плагин в реестре для автозагрузки с Автокад.
        /// </summary>
        [CommandMethod("AutocadUtils_Register")]
        public void RegisterMyApp()
        {
            // Get the AutoCAD Applications key
            string sProdKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            string sAppName = "ArmTools_1.0";

            Microsoft.Win32.RegistryKey regAcadProdKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(sProdKey);
            Microsoft.Win32.RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);

            // Check to see if the "MyApp" key exists
            string[] subKeys = regAcadAppKey.GetSubKeyNames();
            foreach (string subKey in subKeys)
            {
                // If the application is already registered, exit
                if (subKey.Equals(sAppName))
                {
                    regAcadAppKey.Close();
                    return;
                }
            }

            // Get the location of this module
            string sAssemblyPath = Assembly.GetExecutingAssembly().Location;

            // Register the application
            Microsoft.Win32.RegistryKey regAppAddInKey = regAcadAppKey.CreateSubKey(sAppName);
            regAppAddInKey.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String);
            regAppAddInKey.SetValue("LOADCTRLS", 14, RegistryValueKind.DWord);
            regAppAddInKey.SetValue("LOADER", sAssemblyPath, RegistryValueKind.String);
            regAppAddInKey.SetValue("MANAGED", 1, RegistryValueKind.DWord);

            regAcadAppKey.Close();
        }




        /// <summary>
        /// Удаляет регистрацию плагина из реестра.(Убирает автозагрузку в Автокад)  
        /// </summary>
        [CommandMethod("AutocadUtils_Unregister")]
        public void UnregisterMyApp()
        {
            // Get the AutoCAD Applications key
            string sProdKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            string sAppName = "ArmTools_1.0";

            Microsoft.Win32.RegistryKey regAcadProdKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(sProdKey);
            Microsoft.Win32.RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);

            // Delete the key for the application
            regAcadAppKey.DeleteSubKeyTree(sAppName);
            regAcadAppKey.Close();
        }



        public void Initialize()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage(Environment.NewLine + "Пакет утилит. Автор: Семенов Вадим Сергеевич. 2020 год. Email: vad.s.semenov@gmail.com" + Environment.NewLine);
            Autodesk.Windows.ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
        }



         /// <summary>
         /// Обработчик события
         /// Следит за событиями изменения окна автокада.
         /// Используем его для того, чтобы "поймать" момент построения ленты,
         /// учитывая, что наш плагин уже инициализировался
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
        void ComponentManager_ItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
        {
            // Проверяем, что лента загружена
            if (Autodesk.Windows.ComponentManager.Ribbon != null)
            {
                // Строим нашу вкладку
                BuildRibbonTab();

                // Отключаем обработчик событий
                Autodesk.Windows.ComponentManager.ItemInitialized -= ComponentManager_ItemInitialized;
            }
        }

        public void Terminate()
        {
          
        }
    }
}
