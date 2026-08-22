using Infragistics.Windows;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using SasControls;
using SasFormBrowes;
using SasLib;
using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace SasVoucherLib
{
    /// <summary>Lớp thao tác với trường tự do.</summary>
    public class FreeCodeFieldLib
    {
        public static void InitFreeCodeField(DataGridView oGrid, string Ma_ct)
        {
        }

        /// <summary>Gắn thêm trường tự do cho Grid</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="oGrid">Grid cần gắn trường tự do.</param>
        /// <param name="Ma_ct">Mã chứng từ dùng để xác định trường tự do cần thêm.</param>
        /// <param name="Kind">1- ct,2- ctgt</param>
        public static void InitFreeCodeField(
          SasObject SasObj,
          BasicGridView oGrid,
          string Ma_ct,
          int Kind)
        {
            DataTable freeCodeInfo = SasObj.GetFreeCodeInfo(Ma_ct);
            freeCodeInfo.DefaultView.Sort = "order";
            string str1 = SasObj.GetOption("M_LAN").ToString();
            if (Kind == 1 && oGrid.GetType().Name.Equals("DataGridView"))
            {
                if ((oGrid as DataGridView).FormParent != null && (oGrid as DataGridView).FormParent.GetType().BaseType.Name.Equals("FormTrans"))
                    ((oGrid as DataGridView).FormParent as FormTrans).GridCt.Add(oGrid);
            }
            else if (Kind == 2 && oGrid.GetType().Name.Equals("DataGridView") && ((oGrid as DataGridView).FormParent != null && (oGrid as DataGridView).FormParent.GetType().BaseType.Name.Equals("FormTrans")))
                ((oGrid as DataGridView).FormParent as FormTrans).GridCtgt.Add(oGrid);
            foreach (DataRowView dataRowView in freeCodeInfo.DefaultView)
            {
                string str2 = str1 == "V" ? dataRowView["caption"].ToString() : dataRowView["caption2"].ToString();
                int result1 = 100;
                int.TryParse(dataRowView["width"].ToString(), out result1);
                bool result2;
                bool.TryParse(Kind == 1 ? dataRowView["status"].ToString() : dataRowView["status2"].ToString(), out result2);
                int result3;
                int.TryParse(Kind == 1 ? dataRowView["order"].ToString() : dataRowView["order2"].ToString(), out result3);
                if (dataRowView["enabledstatus"].ToString().Contains(Kind.ToString()) && result2)
                {
                    string str3 = "";
                    string strBrowse = result3 < 0 || string.IsNullOrEmpty(str2.Trim()) ? str3 + string.Format("{0}:{1}:h={2}:E:IV", Kind.Equals(1) ? (object)dataRowView["field_ct"].ToString().Trim() : (object)dataRowView["field_gt"].ToString().Trim(), (object)result1, (object)str2) : str3 + string.Format("{0}:{1}:h={2}:E", Kind.Equals(1) ? (object)dataRowView["field_ct"].ToString().Trim() : (object)dataRowView["field_gt"].ToString().Trim(), (object)result1, (object)str2);
                    Field field = SysFunc.CreateFieldList(SasObj, oGrid, strBrowse, (DataTable)null).First<Field>();
                    field.Tag = (object)"NT";
                    if (dataRowView["loai_dmctct"].ToString().Equals("1"))
                    {
                        if (oGrid.GetType().Equals(typeof(DataGridView)))
                        {
                            FrameworkElementFactory frameworkElementFactory = new FrameworkElementFactory(typeof(AutoCompleteTextBox));
                            frameworkElementFactory.SetValue(AutoCompleteTextBox.ListIDProperty, (object)dataRowView["ma_dm"].ToString().Trim());
                            frameworkElementFactory.SetValue(AutoCompleteTextBox.AllowEmtyProperty, (object)true);
                            frameworkElementFactory.SetValue(AutoCompleteTextBox.AllowWrongProperty, (object)false);
                            frameworkElementFactory.SetValue(AutoCompleteTextBox.IsBorderTextBoxProperty, (object)false);
                            BindingEx bindingEx = new BindingEx("Value");
                            bindingEx.RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(CellValuePresenter), 1);
                            frameworkElementFactory.SetBinding(AutoCompleteTextBox.TextProperty, (BindingBase)bindingEx);
                            frameworkElementFactory.SetBinding(AutoCompleteTextBox.ParentControlProperty, (BindingBase)new Binding()
                            {
                                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(Window), 1),
                                Mode = BindingMode.OneWay
                            });
                            frameworkElementFactory.SetBinding(AutoCompleteTextBox.SasObjProperty, (BindingBase)new Binding("BindingSasObj")
                            {
                                Mode = BindingMode.OneWay,
                                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(FormTrans), 1)
                            });
                            BoolInverter boolInverter = new BoolInverter();
                            frameworkElementFactory.SetBinding(AutoCompleteTextBox.IsReadOnlyProperty, (BindingBase)new Binding("DataPresenter.Tag")
                            {
                                Mode = BindingMode.OneWay,
                                Converter = (IValueConverter)boolInverter,
                                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(CellValuePresenter), 1)
                            });
                            DependencyObject ancestorFromType = Utilities.GetAncestorFromType((DependencyObject)oGrid, typeof(FormTrans), true);
                            if (ancestorFromType != null)
                            {
                                FormTrans formTrans = ancestorFromType as FormTrans;
                                if (formTrans.Resources[(object)"AutoCompleteTextBoxEditModeEnable"] != null)
                                    frameworkElementFactory.SetValue(FrameworkElement.StyleProperty, (object)(formTrans.Resources[(object)"AutoCompleteTextBoxEditModeEnable"] as Style));
                            }
                            DataTemplate dataTemplate = new DataTemplate();
                            dataTemplate.VisualTree = frameworkElementFactory;
                            Style style = new Style(typeof(ControlHostEditor));
                            style.BasedOn = oGrid.FindResource((object)"AutoCompleteTextBoxEditor") as Style;
                            Setter setter = new Setter(ControlHostEditor.EditDataTemplateProperty, (object)dataTemplate);
                            style.Setters.Add((SetterBase)setter);
                            field.Settings.EditorStyle = style;
                            field.Settings.AllowFixing = AllowFieldFixing.No;
                            field.Width = new FieldLength?(new FieldLength((double)result1));
                        }
                    }
                    else if (dataRowView["loai_dmctct"].ToString().Equals("2"))
                    {
                        field.Settings.EditAsType = typeof(string);
                        field.Settings.AllowFixing = AllowFieldFixing.No;
                        field.Width = new FieldLength?(new FieldLength((double)result1));
                    }
                    else if (dataRowView["loai_dmctct"].ToString().Equals("3"))
                    {
                        Style style = new Style(typeof(NumericTextBox));
                        Setter setter1 = new Setter(ValueEditor.FormatProperty, (object)SasObj.GetOption("M_IP_SL").ToString());
                        style.Setters.Add((SetterBase)setter1);
                        Setter setter2 = new Setter(TextEditorBase.NullTextProperty, (object)string.Format("{0:" + SasObj.GetOption("M_IP_SL").ToString() + "}", (object)0));
                        style.Setters.Add((SetterBase)setter2);
                        field.Settings.EditorStyle = style;
                        field.Settings.AllowFixing = AllowFieldFixing.No;
                        field.Width = new FieldLength?(new FieldLength((double)result1));
                    }
                    if (field != null)
                    {
                        BindingOperations.SetBinding((DependencyObject)field.Settings, FieldSettings.AllowEditProperty, (BindingBase)new Binding("Tag")
                        {
                            Source = (object)oGrid,
                            Mode = BindingMode.OneWay
                        });
                        oGrid.DefaultFieldLayout.Fields.Insert(Math.Min(result3, oGrid.DefaultFieldLayout.Fields.Count), field);
                    }
                }
            }
        }

        public static void InitFreeCodeField(
          SasObject SasObj,
          BasicGridView oGrid,
          string Ma_ct,
          int Kind,
          DependencyProperty[] props,
          object[] values)
        {
            DataTable freeCodeInfo = SasObj.GetFreeCodeInfo(Ma_ct);
            freeCodeInfo.DefaultView.Sort = "order";
            string str1 = SasObj.GetOption("M_LAN").ToString();
            if (Kind == 1 && oGrid.GetType().Name.Equals("DataGridView"))
            {
                if ((oGrid as DataGridView).FormParent != null && (oGrid as DataGridView).FormParent.GetType().BaseType.Name.Equals("FormTrans"))
                    ((oGrid as DataGridView).FormParent as FormTrans).GridCt.Add(oGrid);
            }
            else if (Kind == 2 && oGrid.GetType().Name.Equals("DataGridView") && ((oGrid as DataGridView).FormParent != null && (oGrid as DataGridView).FormParent.GetType().BaseType.Name.Equals("FormTrans")))
                ((oGrid as DataGridView).FormParent as FormTrans).GridCtgt.Add(oGrid);
            foreach (DataRowView dataRowView in freeCodeInfo.DefaultView)
            {
                string str2 = str1 == "V" ? dataRowView["caption"].ToString() : dataRowView["caption2"].ToString();
                int result1 = 100;
                int.TryParse(dataRowView["width"].ToString(), out result1);
                bool result2;
                bool.TryParse(Kind == 1 ? dataRowView["status"].ToString() : dataRowView["status2"].ToString(), out result2);
                int result3;
                int.TryParse(Kind == 1 ? dataRowView["order"].ToString() : dataRowView["order2"].ToString(), out result3);
                if (dataRowView["enabledstatus"].ToString().Contains(Kind.ToString()))
                {
                    Field field = (Field)null;
                    string str3 = "";
                    string strBrowse = result3 < 0 || !result2 || string.IsNullOrEmpty(str2.Trim()) ? str3 + string.Format("{0}:{1}:h={2}:E:IV", Kind.Equals(1) ? (object)dataRowView["field_ct"].ToString().Trim() : (object)dataRowView["field_gt"].ToString().Trim(), (object)result1, (object)str2) : str3 + string.Format("{0}:{1}:h={2}:E", Kind.Equals(1) ? (object)dataRowView["field_ct"].ToString().Trim() : (object)dataRowView["field_gt"].ToString().Trim(), (object)result1, (object)str2);
                    if (dataRowView["loai_dmctct"].ToString().Equals("1"))
                    {
                        field = SysFunc.CreateFieldList(SasObj, oGrid, strBrowse, (DataTable)null).First<Field>();
                        field.Tag = (object)"NT";
                        FrameworkElementFactory frameworkElementFactory = new FrameworkElementFactory(typeof(AutoCompleteTextBox));
                        frameworkElementFactory.SetValue(AutoCompleteTextBox.ListIDProperty, (object)dataRowView["ma_dm"].ToString().Trim());
                        frameworkElementFactory.SetValue(AutoCompleteTextBox.AllowEmtyProperty, (object)true);
                        frameworkElementFactory.SetValue(AutoCompleteTextBox.AllowWrongProperty, (object)false);
                        frameworkElementFactory.SetValue(AutoCompleteTextBox.IsBorderTextBoxProperty, (object)false);
                        if (props.Length == values.Length)
                        {
                            for (int index = 0; index < props.Length; ++index)
                                frameworkElementFactory.SetValue(props[index], values[index]);
                        }
                        BindingEx bindingEx = new BindingEx("Value");
                        bindingEx.RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(CellValuePresenter), 1);
                        frameworkElementFactory.SetBinding(AutoCompleteTextBox.TextProperty, (BindingBase)bindingEx);
                        frameworkElementFactory.SetBinding(AutoCompleteTextBox.ParentControlProperty, (BindingBase)new Binding()
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(Window), 1),
                            Mode = BindingMode.OneWay
                        });
                        frameworkElementFactory.SetBinding(AutoCompleteTextBox.SasObjProperty, (BindingBase)new Binding("BindingSasObj")
                        {
                            Mode = BindingMode.OneWay,
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(FormTrans), 1)
                        });
                        DependencyObject ancestorFromType = Utilities.GetAncestorFromType((DependencyObject)oGrid, typeof(FormTrans), true);
                        if (ancestorFromType != null)
                        {
                            FormTrans formTrans = ancestorFromType as FormTrans;
                            if (formTrans.Resources[(object)"AutoCompleteTextBoxEditModeEnable"] != null)
                                frameworkElementFactory.SetValue(FrameworkElement.StyleProperty, (object)(formTrans.Resources[(object)"AutoCompleteTextBoxEditModeEnable"] as Style));
                        }
                        DataTemplate dataTemplate = new DataTemplate();
                        dataTemplate.VisualTree = frameworkElementFactory;
                        Style style = new Style(typeof(ControlHostEditor));
                        style.BasedOn = oGrid.FindResource((object)"AutoCompleteTextBoxEditor") as Style;
                        Setter setter = new Setter(ControlHostEditor.EditDataTemplateProperty, (object)dataTemplate);
                        style.Setters.Add((SetterBase)setter);
                        field.Settings.EditorStyle = style;
                        field.Settings.AllowFixing = AllowFieldFixing.No;
                        field.Width = new FieldLength?(new FieldLength((double)result1));
                    }
                    else if (dataRowView["loai_dmctct"].ToString().Equals("2"))
                    {
                        field.Settings.EditAsType = typeof(string);
                        field.Settings.AllowFixing = AllowFieldFixing.No;
                        field.Width = new FieldLength?(new FieldLength((double)result1));
                    }
                    else if (dataRowView["loai_dmctct"].ToString().Equals("3"))
                    {
                        Style style = new Style(typeof(NumericTextBox));
                        Setter setter1 = new Setter(ValueEditor.FormatProperty, (object)SasObj.GetOption("M_IP_SL").ToString());
                        style.Setters.Add((SetterBase)setter1);
                        Setter setter2 = new Setter(TextEditorBase.NullTextProperty, (object)string.Format("{0:" + SasObj.GetOption("M_IP_SL").ToString() + "}", (object)0));
                        style.Setters.Add((SetterBase)setter2);
                        field.Settings.EditorStyle = style;
                        field.Settings.AllowFixing = AllowFieldFixing.No;
                        field.Width = new FieldLength?(new FieldLength((double)result1));
                    }
                    if (field != null)
                    {
                        BindingOperations.SetBinding((DependencyObject)field.Settings, FieldSettings.AllowEditProperty, (BindingBase)new Binding("Tag")
                        {
                            Source = (object)oGrid,
                            Mode = BindingMode.OneWay
                        });
                        oGrid.DefaultFieldLayout.Fields.Insert(Math.Min(result3, oGrid.DefaultFieldLayout.Fields.Count), field);
                    }
                }
            }
        }

        public static void CarryFreeCodeFields(
          SasObject SasObj,
          string Ma_ct,
          DataView CtTableView,
          DataRow newRow,
          int GrdType)
        {
            if (CtTableView.Count <= 0)
                return;
            DataTable freeCodeInfo = SasObj.GetFreeCodeInfo(Ma_ct);
            string index = GrdType == 1 ? "field_ct" : "field_gt";
            foreach (DataRow row in (InternalDataCollectionBase)freeCodeInfo.Rows)
            {
                if (!string.IsNullOrEmpty(row[index].ToString().Trim()) && row["carry"].ToString().Trim() == "1" && CtTableView.ToTable().Columns.Contains(row[index].ToString().Trim()))
                    newRow[row[index].ToString().Trim()] = CtTableView[CtTableView.Count - 1][row[index].ToString().Trim()];
            }
        }

        public static void TransFreeCodeFields(
          SasObject SasObj,
          string Ma_ct,
          DataRow CtRow,
          DataRow GtRow)
        {
            foreach (DataRow row in (InternalDataCollectionBase)SasObj.GetFreeCodeInfo(Ma_ct).Rows)
            {
                if (!string.IsNullOrEmpty(row["field_ct"].ToString()) && !string.IsNullOrEmpty(row["field_gt"].ToString()) && row["carry"].ToString() == "1")
                    GtRow[row["field_gt"].ToString()] = CtRow[row["field_ct"].ToString()];
            }
        }
    }
}
