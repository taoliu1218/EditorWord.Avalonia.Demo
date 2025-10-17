using System;
using Avalonia.Controls;
using EditWord.Avalonia.Until;
using ReactiveUI.SourceGenerators;

namespace EditWord.Avalonia.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{

    [Reactive] private Control? _bodyControl;

    [Reactive] private string _requestA = "我是书签A";
    
    [Reactive] private string _text1 = "这是书签Text1双向绑定内容";
    
    [Reactive] private string _text2 = "这是书签Text2双向绑定内容";
    
    [Reactive] private string _text3 = "这是书签Text3双向绑定内容";
    
    [Reactive] private string _text4 = "这是书签Text4双向绑定内容";

    [ReactiveCommand]
    private void OpenDoc()
    {
        try
        {
            var path = "/Users/ilm/Desktop/test.docx";

            BodyControl = WordReadHelper.RenderDocument(path);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }
        
    }
}