using Consolidate.BizLogic;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace OrderInfoConsolidate
{
    public partial class Ribbon
    {
        private const string originalDefItem = "请选择原始Sheet";
        private const string referenceDefItem = "请选择引用Sheet";

        private Workbook _workBook;
        private Sheets _sheets;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            InitDropDownList();
        }

        #region Start Consolidate
        private void btnStartConsolidate_Click(object sender, RibbonControlEventArgs e)
        {
            if(_sheets == null)
            {
                System.Windows.Forms.MessageBox.Show("请先点击加载Sheet按钮！");
                return;
            }

            string originalSheetName = ddlOriginalSheet.SelectedItem.Label;
            string referenceSheetName = ddlRefSheet.SelectedItem.Label;

            if(originalSheetName == originalDefItem)
            {
                System.Windows.Forms.MessageBox.Show("必须选择一个原数据Sheet！");
            }

            if(referenceSheetName == referenceDefItem)
            {
                System.Windows.Forms.MessageBox.Show("必须选择一个引用数据Sheet！");
            }

            if(originalSheetName == referenceSheetName)
            {
                System.Windows.Forms.MessageBox.Show("原数据Sheet不能和引用数据Sheet是同一个！");
            }

            Worksheet originalWorkSheet = null;
            Worksheet referenceWorkSheet = null;

            foreach(Worksheet sheet in _sheets)
            {
                if(sheet.Name == originalSheetName)
                {
                    originalWorkSheet = sheet;
                }

                if(sheet.Name == referenceSheetName)
                {
                    referenceWorkSheet = sheet;
                }
            }

            if(originalWorkSheet == null)
            {
                System.Windows.Forms.MessageBox.Show("选择的原数据Sheet无效，请检查！");
            }
            if(referenceWorkSheet == null)
            {
                System.Windows.Forms.MessageBox.Show("选择的引用数据Sheet无效，请检查！");
            }
            string errorMsg = string.Empty ;
            ConsolidateBiz biz = new ConsolidateBiz();
            bool result = biz.Consolidate(originalWorkSheet, referenceWorkSheet, out errorMsg);
            if (!result)
            {
                System.Windows.Forms.MessageBox.Show(errorMsg);
            }
            System.Windows.Forms.MessageBox.Show("匹配完成！");
        }
        #endregion

        #region Load Dorp Down List
        private void btnLoadSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _workBook = Globals.Consolidate.Application.ActiveWorkbook;
            if(_workBook == null)
            {
                return;
            }
            _sheets = _workBook.Sheets;
            if(_sheets == null)
            {
                return;
            }

            InitDropDownList();

            foreach (Worksheet workSheet in _sheets)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = workSheet.Name;
                ddlOriginalSheet.Items.Add(item);
            }
            ddlOriginalSheet.SelectedItemIndex = 0;

            
            foreach (Worksheet workSheet in _sheets)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = workSheet.Name;
                ddlRefSheet.Items.Add(item);
            }
            ddlRefSheet.SelectedItemIndex = 0;
        }
        #endregion

        private void InitDropDownList()
        {
            ddlOriginalSheet.Items.Clear();
            RibbonDropDownItem defOriginItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            defOriginItem.Label = originalDefItem;
            ddlOriginalSheet.Items.Add(defOriginItem);

            ddlRefSheet.Items.Clear();
            RibbonDropDownItem defRefItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            defRefItem.Label = referenceDefItem;
            ddlRefSheet.Items.Add(defRefItem);
        }
    }
}
