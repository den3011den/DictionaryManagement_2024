using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;
using Radzen.Blazor;
using System.Text.Json;

namespace DictionaryManagement_Server.Extensions
{
    public class RadzenDataGridApp<TItem> : RadzenDataGrid<TItem>
    {
        [Parameter] public bool? ShowCleanGridSettingsHeaderButton { get; set; } = true;
        [Parameter] public bool? ShowCleanGridFiltersHeaderButton { get; set; } = true;
        [Parameter] public bool? ShowCleanGridSortsHeaderButton { get; set; } = true;
        [Parameter] public string? SettingsName { get; set; } = "";
        [Parameter] public int? MesDepartmentCount { get; set; } = 0;

        [Inject]
        IJSRuntime? JS { get; set; }

        [Inject]
        DialogService _dialogs { get; set; }

        public RadzenDataGridApp() : base()
        {
            base.AndOperatorText = "И";
            base.OrOperatorText = "ИЛИ";
            base.StartsWithText = "Начинается с";
            base.ApplyFilterText = "Применить";
            base.ClearFilterText = "Очистить";
            base.ContainsText = "Содержит";
            base.EndsWithText = "Заканчивается на";
            base.DoesNotContainText = "Не содержит";
            base.EqualsText = "Равно";
            base.GreaterThanOrEqualsText = "Больше или равно";
            base.GreaterThanText = "Больше чем";
            base.IsEmptyText = "Пусто";
            base.IsNotEmptyText = "Не пусто";
            base.LessThanOrEqualsText = "Меньше или равно";
            base.LessThanText = "Меньше чем";
            base.IsNotNullText = "Не пусто";
            base.IsNullText = "Пусто";
            base.NotEqualsText = "Не равно";
            base.FilterText = "Фильтр";
            base.AllColumnsText = "Все";
            base.ColumnsShowingText = "колонок отображается";
            base.ColumnsText = "Не выбрано";
            base.PagingSummaryFormat = $"Страница {{0}} из {{1}} (всего записей {{2}} )";
            base.Density = Density.Compact;
            base.EmptyTemplate = EmptyTemplateRender;
            base.HeaderTemplate = HeaderTemplateRender;
            this.VirtualizationOverscanCount = 50;
        }


        private RenderFragment EmptyTemplateRender => (builder) =>
        {
            builder.OpenElement(1, "p");
            builder.AddAttribute(2, "style", "color: lightgrey; font-size: 24px; text-align: center; margin: 2rem;");
            builder.AddContent(3, "Нет записей для отображения");
            builder.CloseElement();
        };

        private RenderFragment HeaderTemplateRender => (builder) =>
        {
            if (ShowCleanGridSettingsHeaderButton == true || ShowCleanGridFiltersHeaderButton == true || ShowCleanGridSortsHeaderButton == true)
            {
                if (ShowCleanGridSettingsHeaderButton == true)
                {
                    builder.OpenComponent<RadzenButton>(1);
                    builder.AddAttribute(2, "Size", ButtonSize.Small);
                    builder.AddAttribute(3, "Text", "Очистить настройки интерфейса страницы");
                    builder.AddAttribute(4, "Icon", "settings");
                    builder.AddAttribute(5, "ButtonStyle", ButtonStyle.Light);
                    builder.AddAttribute(6, "Variant", Variant.Flat);
                    builder.AddAttribute(7, "Click", Microsoft.AspNetCore.Components.CompilerServices.RuntimeHelpers.TypeCheck<Microsoft.AspNetCore.Components.EventCallback<Microsoft.AspNetCore.Components.Web.MouseEventArgs>>(Microsoft.AspNetCore.Components.EventCallback.Factory.Create<Microsoft.AspNetCore.Components.Web.MouseEventArgs>(this,
                            (args) => CleanPageSettings()
                            )));
                    builder.CloseComponent();
                }
                if (ShowCleanGridFiltersHeaderButton == true)
                {
                    builder.OpenComponent<RadzenButton>(8);
                    builder.AddAttribute(9, "Size", ButtonSize.Small);
                    builder.AddAttribute(10, "Text", "Очистить все фильтры");
                    builder.AddAttribute(11, "Icon", "filter_alt");
                    builder.AddAttribute(12, "ButtonStyle", ButtonStyle.Light);
                    builder.AddAttribute(13, "Variant", Variant.Flat);
                    builder.AddAttribute(14, "Click", Microsoft.AspNetCore.Components.CompilerServices.RuntimeHelpers.TypeCheck<Microsoft.AspNetCore.Components.EventCallback<Microsoft.AspNetCore.Components.Web.MouseEventArgs>>(Microsoft.AspNetCore.Components.EventCallback.Factory.Create<Microsoft.AspNetCore.Components.Web.MouseEventArgs>(this,
                            (args) => CleanAllFilters()
                            )));
                    builder.CloseComponent();
                }

                if (ShowCleanGridSortsHeaderButton == true)
                {
                    builder.OpenComponent<RadzenButton>(15);
                    builder.AddAttribute(16, "Size", ButtonSize.Small);
                    builder.AddAttribute(17, "Text", "Очистить все сортировки");
                    builder.AddAttribute(18, "Icon", "swap_vert");
                    builder.AddAttribute(19, "ButtonStyle", ButtonStyle.Light);
                    builder.AddAttribute(20, "Variant", Variant.Flat);
                    builder.AddAttribute(21, "Click", Microsoft.AspNetCore.Components.CompilerServices.RuntimeHelpers.TypeCheck<Microsoft.AspNetCore.Components.EventCallback<Microsoft.AspNetCore.Components.Web.MouseEventArgs>>(Microsoft.AspNetCore.Components.EventCallback.Factory.Create<Microsoft.AspNetCore.Components.Web.MouseEventArgs>(this,
                            (args) => CleanAllOrders()
                            )));

                    builder.CloseComponent();
                }

                bool addRecordCountSign = true;
                if (this.Data != null)
                    if (this.Data.GetType().ToString().ToUpper().Contains("MESDEPARTMENTDTO"))
                    {
                        addRecordCountSign = false;
                    }
                builder.OpenComponent<RadzenButton>(22);
                builder.AddAttribute(23, "Size", ButtonSize.Small);
                builder.AddAttribute(24, "Style", "text-transform:none;");
                if (addRecordCountSign)
                {
                    builder.AddAttribute(25, "Text", String.Concat("Записей с учётом фильтров ", this.View == null ? "0" : this.View.Count().ToString(),
                        " из ", this.Data == null ? "0" : this.Data.Count().ToString(), " в выборке"));
                }
                else
                {
                    builder.AddAttribute(25, "Text", String.Concat("Отображается ", this.View == null ? "0" : this.View.Count().ToString(),
                        " из ", this.Data == null ? "0" : MesDepartmentCount.ToString(), " записей"));
                }
                builder.AddAttribute(26, "ButtonStyle", ButtonStyle.Primary);
                builder.AddAttribute(27, "Variant", Variant.Text);
                builder.CloseComponent();
            }
            else
            {
                HeaderTemplate = null;
            }
        };

        private async Task CleanPageSettings()
        {
            await Task.CompletedTask;
            var selectionResult = await _dialogs.Confirm("Будут очищены пользовательские настройки страницы: видимость колонок, порядок следования колонок, ширина колонок, применённые фильтры", "Сбросить настройки интерфейса страницы",
                new ConfirmOptions { OkButtonText = "Очистить", CancelButtonText = "Отмена", Left = "30vw" });

            if (selectionResult != true)
            {
                await InvokeAsync(SaveStateAsync);
                return;
            }
            var result = await JS.InvokeAsync<string>("window.localStorage.removeItem", SettingsName);
            //Settings = null;
            //-------------------------
            if (Settings != null)
            {
                foreach (var c in this.Settings.Columns)
                {
                    c.FilterValue = null;
                    c.SecondFilterValue = null;
                }
            }

            //-------------------------

            await InvokeAsync(SaveStateAsync);
            await Task.Delay(200);
            await InvokeAsync(StateHasChanged);

        }

        public async Task CleanAllFilters()
        {
            var selectionResult = await _dialogs.Confirm("Будут очищены все фильтры", "Очистить фильтры",
                new ConfirmOptions { OkButtonText = "Очистить", CancelButtonText = "Отмена", Left = "30vw" });

            if (selectionResult != true)
            {
                await InvokeAsync(SaveStateAsync);
                return;
            }

            if (Settings != null)
            {
                foreach (var c in this.Settings.Columns)
                {
                    c.FilterValue = null;
                    c.SecondFilterValue = null;
                }
            }
            await InvokeAsync(SaveStateAsync);
            await Task.Delay(200);
            await InvokeAsync(StateHasChanged);
        }

        async Task CleanAllOrders()
        {
            var selectionResult = await _dialogs.Confirm("Будут очищены все сортировки", "Очистить сортировки",
                new ConfirmOptions { OkButtonText = "Очистить", CancelButtonText = "Отмена", Left = "30vw" });
            if (selectionResult != true)
            {
                await InvokeAsync(SaveStateAsync);
                return;
            }

            if (Settings != null)
            {
                foreach (var c in Settings.Columns)
                {
                    c.SortOrder = null;
                    c.FilterValue = null;
                    c.SecondFilterValue = null;
                    c.Visible = true;
                    c.OrderIndex = 0;
                    if (this.ColumnWidth != null)
                        c.Width = this.ColumnWidth;
                }
            }
            await InvokeAsync(SaveStateAsync);
            await Task.Delay(200);
            await InvokeAsync(StateHasChanged);
        }

        private async Task SaveStateAsync()
        {
            await Task.CompletedTask;
            await JS.InvokeVoidAsync("eval", "window.localStorage.setItem('" + SettingsName + "', '" + JsonSerializer.Serialize<DataGridSettings>(Settings) + "')");
        }
    }
}