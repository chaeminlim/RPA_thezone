﻿#pragma checksum "..\..\MappingTable.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "371A0CA4B1A1D284CE6E427B7D7C25DC39EADF67D76871F600786BD0116B0E94"
//------------------------------------------------------------------------------
// <auto-generated>
//     이 코드는 도구를 사용하여 생성되었습니다.
//     런타임 버전:4.0.30319.42000
//
//     파일 내용을 변경하면 잘못된 동작이 발생할 수 있으며, 코드를 다시 생성하면
//     이러한 변경 내용이 손실됩니다.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using tempproj;


namespace tempproj {
    
    
    /// <summary>
    /// MappingTable
    /// </summary>
    public partial class MappingTable : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 43 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox ClientTypeComboBox;
        
        #line default
        #line hidden
        
        
        #line 53 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DataGrid MappingDataGrid;
        
        #line default
        #line hidden
        
        
        #line 66 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label StatusLabel;
        
        #line default
        #line hidden
        
        
        #line 68 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox FromTextBox;
        
        #line default
        #line hidden
        
        
        #line 70 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox ToTextBox;
        
        #line default
        #line hidden
        
        
        #line 72 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox TypeTextBox;
        
        #line default
        #line hidden
        
        
        #line 74 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox TypeNameTextBox;
        
        #line default
        #line hidden
        
        
        #line 76 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox TrueTextBox;
        
        #line default
        #line hidden
        
        
        #line 78 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox FalseTextBox;
        
        #line default
        #line hidden
        
        
        #line 79 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btn_AddRow;
        
        #line default
        #line hidden
        
        
        #line 80 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btn_deleteRow;
        
        #line default
        #line hidden
        
        
        #line 81 "..\..\MappingTable.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btn_Save;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/tempproj;component/mappingtable.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\MappingTable.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.ClientTypeComboBox = ((System.Windows.Controls.ComboBox)(target));
            
            #line 43 "..\..\MappingTable.xaml"
            this.ClientTypeComboBox.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.ClientTypeComboBox_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 2:
            this.MappingDataGrid = ((System.Windows.Controls.DataGrid)(target));
            return;
            case 3:
            this.StatusLabel = ((System.Windows.Controls.Label)(target));
            return;
            case 4:
            this.FromTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 5:
            this.ToTextBox = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 6:
            this.TypeTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 7:
            this.TypeNameTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 8:
            this.TrueTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 9:
            this.FalseTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 10:
            this.btn_AddRow = ((System.Windows.Controls.Button)(target));
            
            #line 79 "..\..\MappingTable.xaml"
            this.btn_AddRow.Click += new System.Windows.RoutedEventHandler(this.btn_AddRow_Click);
            
            #line default
            #line hidden
            return;
            case 11:
            this.btn_deleteRow = ((System.Windows.Controls.Button)(target));
            
            #line 80 "..\..\MappingTable.xaml"
            this.btn_deleteRow.Click += new System.Windows.RoutedEventHandler(this.btn_deleteRow_Click);
            
            #line default
            #line hidden
            return;
            case 12:
            this.btn_Save = ((System.Windows.Controls.Button)(target));
            
            #line 81 "..\..\MappingTable.xaml"
            this.btn_Save.Click += new System.Windows.RoutedEventHandler(this.btn_Save_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

