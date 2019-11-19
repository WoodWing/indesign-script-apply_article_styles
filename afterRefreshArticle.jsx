
//
//  -Smart Connection afterRefreshArticle event-
//
//  callback after Smart Connection Update Content has completed
// 
//  incoming parameters:
//      Core_ID     The object id of the article that was refreshed
//

try {    
    var scriptfile = File(app.scriptPreferences.scriptsFolder.fullName+"/../EnterpriseScripts/apply_article_style.jsx");
    app.doScript(scriptfile, ScriptLanguage.JAVASCRIPT, 
                ['afterRefreshArticle', 'Core_ID', app.scriptArgs.get('Core_ID')], 
                UndoModes.ENTIRE_SCRIPT, "Apply article style...");
    
} catch (err) {
    alert( 'EVENT ERROR '+err);
}