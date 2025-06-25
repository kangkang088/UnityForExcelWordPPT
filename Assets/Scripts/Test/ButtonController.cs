using UnityEngine;
using UnityEngine.UI;

public class ButtonController : MonoBehaviour
{
    public Button wordBtn;
    public Button excelBtn;
    public Button pptBtn;
    public Button pdfBtn;

    private void Start()
    {
        //wordBtn.onClick.AddListener(()=>{
        //    LoadWord.Instance.LoadGo(Application.streamingAssetsPath+"/Word/word.docx");
        //});
        //excelBtn.onClick.AddListener(()=>{
        //    LoadExcel.Instance.LoadGo(Application.streamingAssetsPath+"/Excel/excel.xlsx");
        //});
        //pptBtn.onClick.AddListener(() => {
        //    LoadPPT.Instance.LoadGo(Application.streamingAssetsPath + "/PPT/ppt.pptx");
        //});
        //pdfBtn.onClick.AddListener(() => {
        //    LoadPDF.Instance.LoadGo(Application.streamingAssetsPath + "/PDF/pdf.pdf");
        //});

        wordBtn.onClick.AddListener(()=>{
            LoadWordTest.Instance.LoadWordGo(Application.streamingAssetsPath+"/Word/word.docx");
        });
        excelBtn.onClick.AddListener(()=>{
            LoadExcelTest.Instance.LoadExcelGO(Application.streamingAssetsPath+"/Excel/excel.xlsx");
        });
        pptBtn.onClick.AddListener(() => {
            LoadPPTTest.Instance.LoadPPTGO(Application.streamingAssetsPath + "/PPT/ppt.pptx");
        });
        pdfBtn.onClick.AddListener(() => {
            LoadPDFTest.Instance.LoadPDFGo(Application.streamingAssetsPath + "/PDF/pdf.pdf");
        });
    }
}
