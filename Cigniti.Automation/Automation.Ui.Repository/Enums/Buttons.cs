using System.ComponentModel;

namespace Automation.Ui.Repository.Enums
{

    /// <summary>
    /// Buttons Enum
    /// </summary>

    public enum UIBUTTONS
    {
        [Description("Cancel")]
        Cancel,
        [Description("Continue")]
        Continue,
        [Description("Yes")]
        Yes,
        [Description("No")]
        NO,
        [Description("Continue & Next")]
        ContinueandNext,
        [Description("Save")]
        SAVE,
        [Description("Wet Sign")]
        WetSign,
        [Description("E-Sign Documents")]
        EsignDocuments,
        [Description("PushToDMSButton")]
        Pushtodms,
        [Description("ClosePushDealPopup")]
        ClosePDealPopup
    }
}
