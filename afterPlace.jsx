//
//  -Smart Connection afterPlace event-
//
//  callback after Place Article (or Image) has completed
// 
//  incoming parameters:
//      pageitem    The pageitem id of the frame the object was placed into 
//                 
//

try {
    var scriptfile = File(app.scriptPreferences.scriptsFolder.fullName+"/../EnterpriseScripts/apply_article_style.jsx");
    app.doScript(scriptfile, ScriptLanguage.JAVASCRIPT, 
                ['afterPlace', app.scriptArgs.get('pageitem') ], 
                UndoModes.ENTIRE_SCRIPT, "Apply article style...");
} catch (err) {
    alert( 'EVENT ERROR '+err);
}