using System;
using System.Collections.Generic;
using System.Linq;
using Consolidate.Entities;
using Microsoft.Office.Interop.Excel;

namespace Consolidate.BizLogic
{
    /// <summary>
    /// Consolidate Business logic
    /// Author: Leo Wu
    /// Create Date: 2017-3-30
    /// Descript:
    /// 1. Get Reference sheet Order Id, Vendor and Amount columns index.
    /// 2. Get Original sheet Order Id column index.
    /// 3. Read Reference sheet, generate and filling Reference Data list.
    /// 4. Group by Vendor Name attribute in Reference Data list. then generate and filling Vendor Data list.
    /// 5. Add Vendor name column in Original sheet by Vendor Data list.
    /// 6. Consolidate Original sheet data and Reference sheet data, use Order Id is Key column then filling Vendor name column.
    /// 7. All have consolidated Reference Data list data, change HaveConsolidate column to current datetime.
    /// 8. Write back Reference Have Consoliadate column data to Reference sheet.
    /// </summary>
    public class ConsolidateBiz
    {
        private int _orderIdColumnIndex = 0;
        private int _vendorColumnIndex = 0;
        private int _amountColumnIndex = 0;
        private int _referenceLastColumnIndex = 0;
        private int _originalOrderIdColumnIndex = 0;
        private int _originalLastColumnIndex = 0;
        private List<ReferenceData> _referenceDataList;
        private List<VendorInfo> _vendorInfoList;

        #region Consolidate Original sheet and Reference sheet main method, trigger other business logic method 
        public bool Consolidate(Worksheet original,Worksheet reference,out string errorMsg)
        {
            errorMsg = string.Empty;
            try
            {
                // Get Reference Columns Index
                GetReferenceColumnIndex(reference, out errorMsg);
                if (!string.IsNullOrWhiteSpace(errorMsg))
                {
                    return false;
                }
                // Get Original Columns Index
                GetOriginalColumnIndex(original, out errorMsg);
                if (!string.IsNullOrWhiteSpace(errorMsg))
                {
                    return false;
                }
                //Filling Entities
                FillingEntities(reference, out errorMsg);
                if (!string.IsNullOrWhiteSpace(errorMsg))
                {
                    return false;
                }
                //Transmit Vendor name to colums name
                TransmitVendorToColumns(original, out errorMsg);
                if (!string.IsNullOrWhiteSpace(errorMsg))
                {
                    return false;
                }
                //Consolidate and filling data
                TransmitAndFillingData(original, out errorMsg);
                if (!string.IsNullOrWhiteSpace(errorMsg))
                {
                    return false;
                }
                //Write back flag
                WriteBackHaveConsolidateFlag(reference);
            }
            catch(Exception ex)
            {
                errorMsg = ex.Message;
                return false;
            }
            
            return true;
        }
        #endregion

        #region Get Reference Column Index
        private void GetReferenceColumnIndex(Worksheet reference, out string errorMsg)
        {
            errorMsg = string.Empty;
            bool loop = true;
            int columnIndex = 1;
            while (loop)
            {
                object cellValue = reference.Cells[1, columnIndex].Value;

                if (cellValue == null)
                {
                    _referenceLastColumnIndex = columnIndex;
                    loop = false;
                }
                else
                {
                    string value = cellValue.ToString();
                    switch (value)
                    {
                        case Library.ReferenceOrderIdColumn:
                            {
                                _orderIdColumnIndex = columnIndex;
                                break;
                            }
                        case Library.ReferenceVendorColumn:
                            {
                                _vendorColumnIndex = columnIndex;
                                break;
                            }
                        case Library.ReferenceAmountColumn:
                            {
                                _amountColumnIndex = columnIndex;
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                    columnIndex++;
                }
            }

            if (_orderIdColumnIndex == 0)
            {
                errorMsg = $"引用数据Sheet中{Library.ReferenceOrderIdColumn}列未找到，请检查列名！";
            }
            if (_vendorColumnIndex == 0)
            {
                errorMsg = $"引用数据Sheet中{Library.ReferenceVendorColumn}列未找到，请检查列名！";
            }
            if (_amountColumnIndex == 0)
            {
                errorMsg = $"引用数据Sheet中{Library.ReferenceAmountColumn}列未找到，请检查列名！";
            }
            if(_referenceLastColumnIndex == 0)
            {
                errorMsg = "引用数据Sheet中没有任何列，请检查！";
            }
        }
        #endregion

        #region Get Original Column Index
        public void GetOriginalColumnIndex(Worksheet original, out string errorMsg)
        {
            errorMsg = string.Empty;
            bool loop = true;
            int columnIndex = 1;
            while (loop)
            {
                object cellValue = original.Cells[1, columnIndex].Value;

                if (cellValue == null)
                {
                    loop = false;
                    _originalLastColumnIndex = columnIndex;
                }
                else
                {
                    string value = cellValue.ToString();
                    if (value == Library.OriginalOrderIdColumn)
                    {
                        _originalOrderIdColumnIndex = columnIndex;
                    }
                    columnIndex++;
                }
            }
            if(_originalOrderIdColumnIndex == 0)
            {
                errorMsg = $"原数据Sheet中{Library.OriginalOrderIdColumn}列未找到，请检查列名！";
            }
            if(_originalLastColumnIndex == 0)
            {
                errorMsg = "原数据Sheet中没有任何列，请检查！";
            }
        }
        #endregion

        #region Read Reference Sheet, Filling Entities
        private void FillingEntities(Worksheet reference, out string errorMsg)
        {
            errorMsg = string.Empty;
            _referenceDataList = new List<ReferenceData>();
            int rowIndex = 2;  //Row 1 is title, so start from row 2
            bool loop = true;
            while (loop)
            {
                object orderIdCellValue = reference.Cells[rowIndex, _orderIdColumnIndex].Value;
                object vendorCellValue = reference.Cells[rowIndex, _vendorColumnIndex].Value;
                object amountCellValue = reference.Cells[rowIndex, _amountColumnIndex].Value;

                if(orderIdCellValue == null && vendorCellValue ==null && amountCellValue == null)
                {
                    loop = false;
                }
                else
                {
                    if (orderIdCellValue == null)
                    {
                        errorMsg = $@"第{rowIndex}行数据，{Library.ReferenceOrderIdColumn}数据有问题，请检查！";
                        return;
                    }
                    if (vendorCellValue == null)
                    {
                        errorMsg = $@"第{rowIndex}行数据，{Library.ReferenceVendorColumn}数据有问题，请检查！";
                        return;
                    }
                    if (amountCellValue == null)
                    {
                        errorMsg = $@"第{rowIndex}行数据，{Library.ReferenceAmountColumn}数据有问题，请检查！";
                        return;
                    }

                    decimal amountValue;
                    if (!decimal.TryParse(amountCellValue.ToString(), out amountValue))
                    {
                        errorMsg = $@"第{rowIndex}行数据，{Library.ReferenceAmountColumn}数据不是金额，请检查！";
                        return;
                    }

                    ReferenceData data = new ReferenceData
                    {
                        OrderId = orderIdCellValue.ToString(),
                        VendorName = vendorCellValue.ToString(),
                        Amount = amountValue
                    };
                    _referenceDataList.Add(data);
                    rowIndex++;
                }
            }
        }
        #endregion

        #region  Transmit Vendor to Original Columns
        private void TransmitVendorToColumns(Worksheet original, out string errorMsg)
        {
            errorMsg = string.Empty;
            _vendorInfoList = new List<VendorInfo>();
            IEnumerable<IGrouping<string,decimal>> query = _referenceDataList.GroupBy(e=>e.VendorName,e => e.Amount);
            foreach (IGrouping<string, decimal> vendorGroup in query)
            {
                VendorInfo vendor = new VendorInfo
                {
                    VendorName = vendorGroup.Key,
                    ColumnIndex = 0
                };
                _vendorInfoList.Add(vendor);
            }

            if(_vendorInfoList.Count == 0)
            {
                errorMsg = $"合并{Library.ReferenceVendorColumn}数据出错！";
            }

            int rowNumber = 1;
            int columnIndex = _originalLastColumnIndex;
            foreach(VendorInfo vendor in _vendorInfoList)
            {
                original.Cells[rowNumber, columnIndex].Value = vendor.VendorName;
                vendor.ColumnIndex = columnIndex;
                columnIndex++;
            }
        }
        #endregion

        #region Consolidate and Filling Data
        private void TransmitAndFillingData(Worksheet original, out string errorMsg)
        {
            errorMsg = string.Empty;

            int rowIndex = 2;
            bool loop = true;
            DateTime flagDateTime = DateTime.Now;
            while (loop)
            {
                object orderIdCellValue = original.Cells[rowIndex, _originalOrderIdColumnIndex].Value;
                if(orderIdCellValue == null)
                {
                    loop = false;
                }
                else
                {
                    List<ReferenceData> queryResult = _referenceDataList.Where(e => e.OrderId == orderIdCellValue.ToString()).ToList();
                    if(queryResult.Count > 0)
                    {
                        foreach(VendorInfo vendor in _vendorInfoList)
                        {
                            List<ReferenceData> vendorData = queryResult.Where(e => e.VendorName == vendor.VendorName).ToList();
                            if(vendorData.Count > 0)
                            {
                                decimal amount = 0;
                                foreach(ReferenceData data in vendorData)
                                {
                                    amount += data.Amount;
                                    data.HaveConsolidate = flagDateTime.ToShortDateString() + " " + flagDateTime.ToShortTimeString();
                                }
                                original.Cells[rowIndex, vendor.ColumnIndex].Value = amount;
                            }
                            else
                            {
                                original.Cells[rowIndex, vendor.ColumnIndex].Value = 0;
                            }
                        }
                    }
                    rowIndex++;
                }
            }
        }
        #endregion

        #region Write back Reference data have consolidate flag
        private void WriteBackHaveConsolidateFlag(Worksheet reference)
        {
            int rowIndex = 2;  //Row 1 is title, so start from row 2
            bool loop = true;
            while (loop)
            {
                object orderIdCellValue = reference.Cells[rowIndex, _orderIdColumnIndex].Value;
                object vendorCellValue = reference.Cells[rowIndex, _vendorColumnIndex].Value;
                object amountCellValue = reference.Cells[rowIndex, _amountColumnIndex].Value;

                if (orderIdCellValue == null && vendorCellValue == null && amountCellValue == null)
                {
                    loop = false;
                }
                else
                {
                    decimal amountValue = decimal.Parse(amountCellValue.ToString());
                    ReferenceData data = _referenceDataList.FirstOrDefault(e => e.OrderId == orderIdCellValue.ToString() && e.VendorName == vendorCellValue.ToString() && e.Amount == amountValue);
                    if (data != null)
                    {
                        reference.Cells[rowIndex, _referenceLastColumnIndex].Value = data.HaveConsolidate;
                    }
                    rowIndex++;
                }
            }
        }
        #endregion
    }
}
