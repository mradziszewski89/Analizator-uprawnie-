// ============================================================
// SharePoint Permission Analyzer - Data File
// Wygenerowany: 2026-04-15 09:49:40
// NIE EDYTUJ TEGO PLIKU RECZNIE
// ============================================================

(function() {
    'use strict';

    // Dane zakodowane w Base64 dla bezpieczenstwa (unikanie problemow z escapowaniem)
    var _b64 = "eyJTY2FuTWV0YWRhdGEiOnsiU2NhblNlc3Npb25JZCI6IkY3Q0Q0ODEwIiwiU2NhblN0YXJ0VGltZSI6IjIwMjYtMDQtMTVUMDk6NDk6MzcuMTg0NzU2MyswMDowMCIsIlNjYW5FbmRUaW1lIjoiMjAyNi0wNC0xNVQwOTo0OTo0MC4yOTQwNDQzKzAwOjAwIiwiU2NhbkR1cmF0aW9uIjo4LCJGYXJtTmFtZSI6IlNoYXJlUG9pbnRfQ29uZmlnIiwiRmFybUJ1aWxkIjoiMTYuMC4xOTEyNy4yMDMzOCIsIlNjYW5uZXJWZXJzaW9uIjoiMS4wLjAiLCJTY2FuU2VydmVyIjoiTU9TU1dGRVNFIiwiU2NhblVzZXIiOiJURVNUXFxsYWIubS5yYWR6aXN6ZXdza2kiLCJDb25maWciOnsiRXhwYW5kU1BHcm91cHMiOnRydWUsIlJhd0Fzc2lnbm1lbnRzT25seSI6ZmFsc2UsIkV4cGFuZERvbWFpbkdyb3VwcyI6dHJ1ZSwiU2NhbkRlcHRoIjp7IlNjYW5XZWJBcHBsaWNhdGlvbnMiOnRydWUsIlNjYW5TaXRlQ29sbGVjdGlvbnMiOnRydWUsIlNjYW5XZWJzIjp0cnVlLCJTY2FuTGlzdHMiOnRydWUsIlNjYW5MaWJyYXJpZXMiOnRydWUsIlNjYW5Gb2xkZXJzIjp0cnVlLCJTY2FuRmlsZXMiOnRydWUsIlNjYW5MaXN0SXRlbXMiOnRydWUsIk1heEl0ZW1zUGVyTGlzdCI6MCwiX2NvbW1lbnRfTWF4SXRlbXNQZXJMaXN0IjoiMCA9IGJleiBsaW1pdHUuIFVzdGF3IG5wLiA1MDAwIGRsYSBzenlic3plZ28gcHJ6ZWJpZWd1IHRlc3Rvd2VnbyJ9fSwiTG9nRmlsZVBhdGgiOiJFOlxcQWRtaW5EYXRhXFxBbmFsaXphdG9yIHVwcmF3bmllxYRcXEFuYWxpemF0b3IgdXByYXduaWXFhFxcUG93ZXJTaGVsbFxcTG9nc1xcU2Nhbl8yMDI2LTA0LTE1XzA5LTQ5LTMxLmxvZyJ9LCJTdGF0aXN0aWNzIjp7IldlYkFwcGxpY2F0aW9uQ291bnQiOjAsIlNpdGVDb2xsZWN0aW9uQ291bnQiOjAsIldlYkNvdW50IjowLCJMaXN0Q291bnQiOjAsIkZvbGRlckNvdW50IjowLCJJdGVtQ291bnQiOjAsIlVuaXF1ZVBlcm1pc3Npb25zQ291bnQiOjAsIlRvdGFsQXNzaWdubWVudHMiOjAsIlRvdGFsT2JqZWN0c1NjYW5uZWQiOjAsIlNraXBwZWRPYmplY3RzIjowLCJFcnJvckNvdW50IjowfSwiRXJyb3JzIjpbXSwiT2JqZWN0cyI6W119";

    try {
        var _json = decodeURIComponent(
            Array.prototype.map.call(
                atob(_b64),
                function(c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                }
            ).join('')
        );
        window.SCAN_DATA = JSON.parse(_json);
    } catch(e) {
        console.error('Blad ladowania danych raportu:', e);
        window.SCAN_DATA = null;
        window.SCAN_DATA_ERROR = e.message;
    }

    window.REPORT_TITLE = "Raport Uprawnien SharePoint - SharePoint_Config";
    window.REPORT_GENERATED = "2026-04-15 09:49:40";
    window.REPORT_SERVER = "MOSSWFESE";

})();
