//
//
//	-apply_article_style-
//
//
//	This script replaces paragraph styles and character styles in all text frames of an article.
//	The replacement styles are organized in	style groups. 
//	The Style Group name is taken from the frame's ObjectStyle name (if applied)
//  or from the (indesign) article name.
//
//	Each (indesign) article/frame on the layout has its unique style group and as such can be
//	individually styled 
//
//	Preferably this script is called automatically after an article is placed or updated.
//	For the demo this script is triggered by 'afterPlace' or by 'afterRefreshArticle'
//



(function apply_article_style() {
	
	try {
		// script operates on active document
		if ("activeDocument" in app)
			var doc = app.activeDocument;
		else
			var doc = app.documents[0];

		// get mode for applying styles:
		// 'all'		-> find replacement for every style
		// 'pseudo'		-> find replacement only for 'pseudo' styles
		var mode = lookup_context();

		// alert(mode);

		// lookup article based on selection or script argument
		var article = lookup_article();
		
		if (article != null) {
			// now update styles on each member of the article
			var textitems = [];
			
			// iterate members to collect text items
			var members = article.articleMembers;
			for (var i=0; i<members.length; i++) {
		
				// article member represents a page item
				var item = members[i].itemRef;
			
				// if the pageitem implements 'parentStory' it is a text frame
				if ('parentStory' in item) {
					textitems.push(item);
				}
				
				// Is item a Group ?  
				if ('ungroup' in item) { 
// 					textitems += item.textFrames;
					for (var j=0; j<item.textFrames.length; j++) {
						textitems.push(item.textFrames[j]);
					}
				}
			}

// 			alert(textitems);
			
			for (var i=0; i<textitems.length; i++) {
				var item = textitems[i]	
				//
				var obs = item.appliedObjectStyle.name;
				if (obs[0] != '[') {
					group = obs;
					if (item.appliedObjectStyle.enableParagraphStyle) {
						// skip this frame, object style will apply the paragraph style(s)
						continue;
					}
				}
				else
					group = article.name;
				
				// alert(group);
				// we need the story, as it contains the text
				var story = item.parentStory;
			
				// locate new style groups (paragraph and character) for the article name
				var pg = doc.paragraphStyleGroups.itemByName(group);
				var cg = doc.characterStyleGroups.itemByName(group);

				// combinations of para styles and char styles are applied on so called
				// 'textStyleRanges'. We will iterate those textStyleRanges, which is
				// faster than iterating Paragraphs and Characters separately
				for (var p=0; p<story.textStyleRanges.length; p++) {
					var tsr = story.textStyleRanges[p];

					// get the style objects from the new group
					// and apply them (ignore error if style does not exist in new group)
					try {
						var ps = lookup_style(pg.paragraphStyles,tsr.appliedParagraphStyle.name, mode);	
						if (ps) tsr.applyParagraphStyle(ps);
					} catch(err) {
// 						tsr.fillColor = 'Cyan';
					}
				
					try {
						var cs = lookup_style(cg.characterStyles,tsr.appliedCharacterStyle.name, mode);
						if (cs) tsr.applyCharacterStyle(cs);
					} catch(err) {
// 						tsr.fillColor = 'Cyan';					
					}
				}
			}
		}

		// 
		// -- helper functions --
		//

		//
		//	- lookup_context -
		//
		function lookup_context() {

			// if (app.selection.length > 0) 
			// 	return 'all';
			
			// if (app.scriptArgs.isDefined('pageitem'))
			// 	return 'all';

			if (app.scriptArgs.isDefined('Core_ID'))
				return 'pseudo';

			return 'all';
		}


		//
		//	- lookup_article -
		//
		//	Find article to apply style on. 
		//	1. selected article
		//	2. 'pageitem' refers to target frame when afterPlace.jsx is triggered
		//	3. 'Core_ID' refers to updated article when afterRefreshArticle.jsx is triggered
		//
		function lookup_article() {
			try {
				if (app.selection.length > 0) {
	 				var item = app.selection[0];
				
					if (!('allArticles' in item)) {
						item = item.parentTextFrames[0];
					}
				}

				if (!item && app.scriptArgs.isDefined('pageitem')) {
					item = app.documents[0].allPageItems.getItemByID(app.scriptArgs.get('pageitem'));
				}
				
				if (!item && app.scriptArgs.isDefined('Core_ID')) {
					var core_id = app.scriptArgs.get('Core_ID');
					var managedarticles = app.documents[0].managedArticles
					for (var i=0; i<managedarticles.length; i++) {
						var managedarticle = managedarticles[i];
						if (managedarticle.entMetaData.get('Core_ID') == core_id) {
							if (managedarticle.components[0].textContainers.length > 0)
								item = managedarticle.components[0].textContainers[0];
						}
					}	
				}

				if (('allArticles' in item)) {
					if (item.allArticles.length) {
						return item.allArticles[0];
					}
				}
						
				return null;
			} 
			catch (err) {
				alert(['lookup_article',err]);
			}
		}

		//
		//	- lookup_style -
		//
		//	Lookup matching style(name) in stylegroup collection
		//
		function lookup_style(styles, stylename, mode) {
			try {
				var stylebase = stylename;

				if (mode == 'all') {
					var stylebase = stylename.split(/[ _-]/);
					if (stylebase.length >= 2) {
						stylebase = stylename.substr(0,9);
					}
					else {
						stylebase = stylebase[0];
					}
				}

				for (var i=0; i<styles.length; i++) {
					if (styles[i].name.indexOf(stylebase) == 0)
						return (styles[i]);
				}
			} catch (err) {
				// do not expect any errors, just in case...
				alert(['lookup_style',err]);
			}
		}
	} catch (err) {
		// do not expect any errors, just in case...
		alert(['apply_article_style',err]);
	}
})();
