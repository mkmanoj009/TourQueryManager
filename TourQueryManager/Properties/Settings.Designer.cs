﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TourQueryManager.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.8.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("server=10.0.2.2;user id=pky;password=pky123;database=tourquerymanagement;SslMode=" +
            "none")]
        public string mysqlConnStr {
            get {
                return ((string)(this["mysqlConnStr"]));
            }
            set {
                this["mysqlConnStr"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("server=10.0.2.2;user id=pky;password=pky123;persistsecurityinfo=True;database=tou" +
            "rquerymanagement;SslMode=none")]
        public string mysqlConnStrMacBook {
            get {
                return ((string)(this["mysqlConnStrMacBook"]));
            }
            set {
                this["mysqlConnStrMacBook"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("server=localhost;user id=pky;password=pky123;persistsecurityinfo=True;database=to" +
            "urquerymanagement;SslMode=none")]
        public string mysqlConnStrLocal {
            get {
                return ((string)(this["mysqlConnStrLocal"]));
            }
            set {
                this["mysqlConnStrLocal"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("server=localhost;user id=excursion;password=excursion123;persistsecurityinfo=True" +
            ";database=tourquerymanagement;SslMode=none")]
        public string mysqlConnStrExcursion {
            get {
                return ((string)(this["mysqlConnStrExcursion"]));
            }
            set {
                this["mysqlConnStrExcursion"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("server=192.168.1.76;user id=pky;password=pky123;persistsecurityinfo=True;database" +
            "=tourquerymanagement;SslMode=none")]
        public string mysqlConnStrUbuntu {
            get {
                return ((string)(this["mysqlConnStrUbuntu"]));
            }
            set {
                this["mysqlConnStrUbuntu"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("server=59.94.224.139;user id=pky;password=pky123;persistsecurityinfo=True;databas" +
            "e=tourquerymanagement;port=55;SslMode=none")]
        public string mysqlConnStrPublic {
            get {
                return ((string)(this["mysqlConnStrPublic"]));
            }
            set {
                this["mysqlConnStrPublic"] = value;
            }
        }
    }
}
