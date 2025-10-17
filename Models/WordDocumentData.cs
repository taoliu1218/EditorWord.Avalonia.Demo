using Avalonia.Controls;

namespace EditWord.Avalonia.Models;

public class WordDocumentData
{
    public Control? HeaderControls { get; set; }

    public Control? BodyControls { get; set; }

    public Control? FooterControls { get; set; }
}