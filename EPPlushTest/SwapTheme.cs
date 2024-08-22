using System.Runtime.Remoting.Contexts;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using AccoRevit.Addin.Application.Services;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

using Nice3point.Revit.Toolkit.External;

using ricaun.Revit.UI;
using ricaun.Revit.UI.Utils;

namespace AccoRevit.Addin.Application.RibbonCommands;

[UsedImplicitly]
[Transaction(TransactionMode.Manual)]
public class SwapTheme : ExternalCommand
{
    public override void Execute()
    {
        try
        {
            UIThemeManager.CurrentTheme = UIThemeManager.CurrentTheme == UITheme.Light ? UITheme.Dark : UITheme.Light;
            SwapIconTheme();
        }
        catch (Exception)
        {
            //message = ex.Message;
        }
    }
    private static void SwapIconTheme()
    {
        List<RibbonItem> allRibbonItems = RibbonController.GetRibbonItems(Context.UiApplication);

        foreach (var ribbonItem in allRibbonItems)
        {
            if (ribbonItem is not PulldownButton)
            {
                var themeImageSource = ribbonItem.GetRibbonItem().Image.GetThemeImageSource();
                ribbonItem?.SetImage(themeImageSource);
                ribbonItem?.SetLargeImage(ribbonItem.GetRibbonItem().LargeImage.GetThemeImageSource());
            }
            else if (ribbonItem is PulldownButton button)
            {
                button?.SetImage(button.Image.GetThemeImageSource());
                button.SetLargeImage(button.LargeImage.GetThemeImageSource());

            }
        }
    }
    public static TRibbonItem SetImage<TRibbonItem>(this TRibbonItem ribbonItem, ImageSource image) where TRibbonItem : Autodesk.Revit.UI.RibbonItem
    {
        image = image.GetThemeImageSource<ImageSource>(RibbonThemeUtils.IsLight);
        Autodesk.Revit.UI.RibbonButton ribbonButton = (object)ribbonItem as Autodesk.Revit.UI.RibbonButton;
        if (ribbonButton != null)
        {
            ribbonButton.Image = image != null ? image.GetBitmapFrame<ImageSource>(16, (Action<ImageSource>)(frame => ribbonButton.Image = frame)) : (ImageSource)null;
        }
        else
        {
            ComboBox comboBox = (object)ribbonItem as ComboBox;
            if (comboBox != null)
            {
                comboBox.Image = image != null ? image.GetBitmapFrame<ImageSource>(16, (Action<ImageSource>)(frame => comboBox.Image = frame)) : (ImageSource)null;
            }
            else
            {
                ComboBoxMember comboBoxMember = (object)ribbonItem as ComboBoxMember;
                if (comboBoxMember != null)
                {
                    comboBoxMember.Image = image != null ? image.GetBitmapFrame<ImageSource>(16, (Action<ImageSource>)(frame => comboBoxMember.Image = frame)) : (ImageSource)null;
                }
                else
                {
                    TextBox textBox = (object)ribbonItem as TextBox;
                    if (textBox != null)
                        textBox.Image = image != null ? image.GetBitmapFrame<ImageSource>(16, (Action<ImageSource>)(frame => textBox.Image = frame)) : (ImageSource)null;
                }
            }
        }
        return ribbonItem;
    }

    public static TRibbonItem SetLargeImage<TRibbonItem>(
      this TRibbonItem ribbonItem,
      ImageSource largeImage)
      where TRibbonItem : Autodesk.Revit.UI.RibbonItem
    {
        largeImage = largeImage.GetThemeImageSource<ImageSource>(RibbonThemeUtils.IsLight);
        Autodesk.Revit.UI.RibbonButton ribbonButton = (object)ribbonItem as Autodesk.Revit.UI.RibbonButton;
        if (ribbonButton != null)
        {
            ribbonButton.LargeImage = largeImage != null ? largeImage.GetBitmapFrame<ImageSource>(32, (Action<ImageSource>)(frame => ribbonButton.LargeImage = frame)) : (ImageSource)null;
            if (ribbonButton.Image == null || ribbonButton.LargeImage == null || ribbonButton.LargeImage is BitmapFrame)
                ribbonButton.SetImage<Autodesk.Revit.UI.RibbonButton>(ribbonButton.LargeImage);
        }
        else
        {
            switch (ribbonItem)
            {
                case ComboBox ribbonItem1:
                    ribbonItem1.SetImage<ComboBox>(largeImage);
                    break;
                case ComboBoxMember ribbonItem2:
                    ribbonItem2.SetImage<ComboBoxMember>(largeImage);
                    break;
                case TextBox ribbonItem3:
                    ribbonItem3.SetImage<TextBox>(largeImage);
                    break;
            }
        }
        return ribbonItem;
    }

    public static Autodesk.Windows.RibbonItem GetRibbonItem(this Autodesk.Revit.UI.RibbonItem ribbonItem)
    {
        return ribbonItem.GetRibbonItem_Alternative();
    }

    public static TImageSource GetThemeImageSource<TImageSource>(
      this TImageSource imageSource,
      bool isLight = true)
      where TImageSource : ImageSource
    {
        string imageTheme;
        return imageSource.GetSourceName().TryThemeImage(isLight, out imageTheme) && imageTheme.GetBitmapSource() is TImageSource bitmapSource ? bitmapSource : imageSource;
    }
}

