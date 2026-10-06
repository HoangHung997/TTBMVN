using System;
using ExcelAddIn1.Funtion;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public partial class Dutoan : Form
    {
        private readonly Excel.Workbook workbook;
        private readonly WorkbookSheetChangeCoordinator coordinator;
        private Form generalForm;
        private NormLookupControl normLookupControl;
        private CostRuleControl costRuleControl;
        private PriceProfileControl priceProfileControl;
        private EstimateRateCatalogControl unitRateControl;
        private EstimateAppendixControl estimateAppendixControl;
        private CostSummaryControl costSummaryControl;
        private ResultAuditControl resultAuditControl;
        private WorkbookValidationControl workbookValidationControl;
        private PackageMigrationControl packageMigrationControl;
        private UpdateCenterControl updateCenterControl;

        public Dutoan()
        {
            InitializeComponent();
            btnMain.Click += (sender, args) => ShowGeneralInformation();
            btnNormLookup.Click += (sender, args) => ShowNormLookup();
            btnCostRule.Click += (sender, args) => ShowCostRule();
            btnPriceProfile.Click += (sender, args) => ShowPriceProfile();
            btnUnitRate.Click += (sender, args) => ShowUnitRate();
            btnEstimateAppendix.Click += (sender, args) => ShowEstimateAppendix();
            btnCostSummary.Click += (sender, args) => ShowCostSummary();
            btnResultAudit.Click += (sender, args) => ShowResultAudit();
            btnWorkbookValidation.Click += (sender, args) => ShowWorkbookValidation();
            btnPackageMigration.Click += (sender, args) => ShowPackageMigration();
            btnUpdateCenter.Click += (sender, args) => ShowUpdateCenter();
            btnSupport.Click += (sender, args) => ShowSupport();
            pnlLeft.Controls.SetChildIndex(btnMain, 0);
            pnlLeft.Controls.SetChildIndex(btnEstimateAppendix, 1);
            pnlLeft.Controls.SetChildIndex(btnPriceProfile, 2);
            pnlLeft.Controls.SetChildIndex(btnUnitRate, 3);
            pnlLeft.Controls.SetChildIndex(btnCostSummary, 4);
            pnlLeft.Controls.SetChildIndex(btnNormLookup, 5);
            pnlLeft.Controls.SetChildIndex(btnCostRule, 6);
            pnlLeft.Controls.SetChildIndex(btnResultAudit, 7);
            pnlLeft.Controls.SetChildIndex(btnWorkbookValidation, 8);
            pnlLeft.Controls.SetChildIndex(btnPackageMigration, 9);
            pnlLeft.Controls.SetChildIndex(btnUpdateCenter, 10);
            pnlLeft.Controls.SetChildIndex(btnSupport, 11);
        }

        public Dutoan(
            Excel.Workbook workbook,
            WorkbookSheetChangeCoordinator coordinator)
            : this(workbook, coordinator, true)
        {
        }

        internal Dutoan(
            Excel.Workbook workbook,
            WorkbookSheetChangeCoordinator coordinator,
            bool showInitialView)
            : this()
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.coordinator = coordinator;
            if (this.coordinator != null)
                this.coordinator.WorkbookClosing += Coordinator_WorkbookClosing;
            Text = "Dự toán rà phá bom mìn";
            if (showInitialView)
                Shown += (sender, args) => ShowGeneralInformation();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (coordinator != null)
                coordinator.WorkbookClosing -= Coordinator_WorkbookClosing;
            normLookupControl?.Dispose();
            normLookupControl = null;
            costRuleControl?.Dispose();
            costRuleControl = null;
            priceProfileControl?.Dispose();
            priceProfileControl = null;
            unitRateControl?.Dispose();
            unitRateControl = null;
            estimateAppendixControl?.Dispose();
            estimateAppendixControl = null;
            costSummaryControl?.Dispose();
            costSummaryControl = null;
            resultAuditControl?.Dispose();
            resultAuditControl = null;
            workbookValidationControl?.Dispose();
            workbookValidationControl = null;
            packageMigrationControl?.Dispose();
            packageMigrationControl = null;
            updateCenterControl?.Dispose();
            updateCenterControl = null;
            generalForm?.Dispose();
            generalForm = null;
            base.OnFormClosed(e);
        }

        private void Coordinator_WorkbookClosing(Excel.Workbook closingWorkbook)
        {
            if (!IsSameWorkbook(workbook, closingWorkbook) || IsDisposed)
                return;
            if (IsHandleCreated)
                BeginInvoke(new Action(Close));
            else
                Close();
        }

        private static bool IsSameWorkbook(Excel.Workbook left, Excel.Workbook right)
        {
            if (left == null || right == null)
                return false;
            if (ReferenceEquals(left, right))
                return true;
            try
            {
                return string.Equals(left.FullName, right.FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private void ShowCostRule()
        {
            try
            {
                if (costRuleControl == null || costRuleControl.IsDisposed)
                {
                    costRuleControl = new CostRuleControl(
                        WorkbookCostRuleService.LoadPinnedCatalog(workbook));
                    pnlRight.Controls.Add(costRuleControl);
                }
                ShowContent(costRuleControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open cost rule");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được bảng chi phí",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowPriceProfile()
        {
            try
            {
                if (priceProfileControl == null || priceProfileControl.IsDisposed)
                {
                    priceProfileControl = new PriceProfileControl(workbook);
                    pnlRight.Controls.Add(priceProfileControl);
                }
                ShowContent(priceProfileControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open price profile");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được bảng giá",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowUnitRate()
        {
            try
            {
                if (unitRateControl == null || unitRateControl.IsDisposed)
                {
                    unitRateControl = new EstimateRateCatalogControl(
                        workbook,
                        ExcelCulture.GetNumberCulture(workbook.Application));
                    pnlRight.Controls.Add(unitRateControl);
                }
                ShowContent(unitRateControl);
                unitRateControl.RefreshRates();
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open unit rate");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được bảng đơn giá",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowEstimateAppendix()
        {
            try
            {
                if (estimateAppendixControl == null || estimateAppendixControl.IsDisposed)
                {
                    estimateAppendixControl = new EstimateAppendixControl(
                        workbook,
                        ExcelCulture.GetNumberCulture(workbook.Application));
                    pnlRight.Controls.Add(estimateAppendixControl);
                }
                ShowContent(estimateAppendixControl);
                estimateAppendixControl.RefreshPreview();
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open estimate appendix");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được phụ lục dự toán",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowCostSummary()
        {
            try
            {
                if (costSummaryControl == null || costSummaryControl.IsDisposed)
                {
                    costSummaryControl = new CostSummaryControl(
                        workbook,
                        ExcelCulture.GetNumberCulture(workbook.Application));
                    pnlRight.Controls.Add(costSummaryControl);
                }
                ShowContent(costSummaryControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open cost summary");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được tổng hợp kinh phí",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowResultAudit()
        {
            try
            {
                if (resultAuditControl == null || resultAuditControl.IsDisposed)
                {
                    resultAuditControl = new ResultAuditControl(workbook);
                    pnlRight.Controls.Add(resultAuditControl);
                }
                ShowContent(resultAuditControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open result audit");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được truy vết căn cứ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowWorkbookValidation()
        {
            try
            {
                if (workbookValidationControl == null || workbookValidationControl.IsDisposed)
                {
                    workbookValidationControl = new WorkbookValidationControl(workbook);
                    pnlRight.Controls.Add(workbookValidationControl);
                }
                ShowContent(workbookValidationControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open workbook validation");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được kiểm tra workbook",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowPackageMigration()
        {
            try
            {
                if (packageMigrationControl == null || packageMigrationControl.IsDisposed)
                {
                    packageMigrationControl = new PackageMigrationControl(workbook);
                    pnlRight.Controls.Add(packageMigrationControl);
                }
                ShowContent(packageMigrationControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open package migration");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được chuyển package",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowUpdateCenter()
        {
            try
            {
                if (updateCenterControl == null || updateCenterControl.IsDisposed)
                {
                    updateCenterControl = new UpdateCenterControl(workbook);
                    updateCenterControl.MigrationRequested += ShowPackageMigrationTarget;
                    pnlRight.Controls.Add(updateCenterControl);
                }
                ShowContent(updateCenterControl);
                updateCenterControl.RefreshInstalled();
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Open update center", "DT-602", "open", workbook, string.Empty);
                MessageBox.Show(
                    ex.Message,
                    "Không mở được cập nhật",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowPackageMigrationTarget(ExcelAddIn1.Core.RegulationPackage package)
        {
            if (package == null)
                return;
            packageMigrationControl?.Dispose();
            packageMigrationControl = null;
            ShowPackageMigration();
            if (packageMigrationControl != null &&
                packageMigrationControl.SelectTarget(package.PackageId, package.DataVersion))
            {
                packageMigrationControl.PreviewSelected();
            }
        }

        private void ShowSupport()
        {
            try
            {
                using (var form = new FrmSupport(workbook))
                    form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Open estimate support", "DT-604", "open", workbook, string.Empty);
                MessageBox.Show(
                    ex.Message,
                    "Không mở được hỗ trợ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowGeneralInformation()
        {
            if (generalForm == null || generalForm.IsDisposed)
            {
                generalForm = new FormThongtinchung
                {
                    TopLevel = false,
                    FormBorderStyle = FormBorderStyle.None,
                    Dock = DockStyle.Fill
                };
                pnlRight.Controls.Add(generalForm);
            }
            ShowContent(generalForm);
            generalForm.Show();
        }

        private void ShowNormLookup()
        {
            try
            {
                if (normLookupControl == null || normLookupControl.IsDisposed)
                {
                    normLookupControl = new NormLookupControl(
                        WorkbookNormCatalogService.LoadPinnedSearchIndex(workbook));
                    pnlRight.Controls.Add(normLookupControl);
                }
                ShowContent(normLookupControl);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open norm lookup");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được tra cứu định mức",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ShowContent(Control selected)
        {
            foreach (Control control in pnlRight.Controls)
                control.Visible = ReferenceEquals(control, selected);
            selected.Visible = true;
            selected.BringToFront();
        }
    }
}
