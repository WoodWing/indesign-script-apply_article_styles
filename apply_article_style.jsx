//
//	This script replaces para styles and char styles in an article based on
//	style groups. The Style Group name is taken from the frame's ObjectStyle name (if applied)
//  ot from the (indesign) article name.
//
//	Each (indesign) article/frame on the layout has it's unique style group and as such can be
//	individually styled 
//
//	Preferably this script is called automatically after an article is placed or updated.
//	For the demo this script is triggered by 'afterPlace' or by 'afterRefreshArticle'
//


//
// TODO - HB
//  
//  - github
//
//	- tijdens place alle stijlen
//	- tijdens refresh alleen pseudostijlen vertalen (ivm nieuwe tussenkoppen etc.)
//	- missende stijl 1 level hoger zoeken
//	- weergave missende stijlen
//	- auto checkout-checkin
//	- enterprise plug-in deployment
//
//	- gedachte: zodra artikel geplaatst is, kan content planner óf een taak in de wcml
//	  evt óók de juiste stijlen gaan gebruiken
//	
//	



(function apply_article_style() {

	//
	//	Find article to apply style on. 
	//	1. selected article
	//	2. 'pageitem' refers to target frame when afterPlace.jsx is triggered
	//	3. 'Core_ID' refers to updated article when afterRefreshArticle.jsx is triggered
	//
	function lookup_article() {
		var item = app.selection[0];
		
		if (!('allArticles' in item)) {
			item = item.parentTextFrames[0];
		}
		
		if (!item && app.scriptArgs.isDefined('pageitem')) {
			var item = app.documents[0].allPageItems.getItemByID(app.scriptArgs.get('pageitem'));
			// alert(item);
		}
		
		if (!item && app.scriptArgs.isDefined('Core_ID')) {
			var core_id = app.scriptArgs.get('Core_ID');
			var managedarticles = app.documents[0].managedArticles
			for (var i=0; i<managedarticles.length; i++) {
				var managedarticle = managedarticles[i];
				if (managedarticle.entMetaData.get('Core_ID') == core_id) {
					var item = managedarticle.components[0].textFrames[0];
				}
			}	
			// alert(item);
		}

		if (('allArticles' in item)) {
			if (item.allArticles.length) {
				return item.allArticles[0];
			}
		}
				
		return null;
	}

	//
	//	Lookup matching style(name) in stylegroup collection
	//
	function lookup_style(styles, stylename) {
 		var stylebase = stylename.split(/[ _-]/);
 		if (stylebase.length > 2) {
			stylebase = stylename.substr(0,9);
		}
		else {
			stylebase = stylebase[0];
		}

		for (var i=0; i<styles.length; i++) {
			if (styles[i].name.indexOf(stylebase) == 0)
				return (styles[i]);
		}
	}

	
	try {
		// script operates on selected item
		var doc = app.activeDocument;
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
						// skip this frame, object style will do the job
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
				
					// get para style name and char style name
					var psname  = tsr.appliedParagraphStyle.name;
					var csname  = tsr.appliedCharacterStyle.name;
			
					// get the style objects from the new group
					// and apply them (ignore error if style does not exist in new group)
					try {
						var ps = lookup_style(pg.paragraphStyles,psname);	
						tsr.applyParagraphStyle(ps);
					} catch(err) {
// 						tsr.fillColor = 'Cyan';
					}
				
					try {
						var cs = lookup_style(cg.characterStyles,csname);
						tsr.applyCharacterStyle(cs);
					} catch(err) {
// 						tsr.fillColor = 'Cyan';					
					}
				}
			}
		}
	} catch (err) {
		// do not expect any errors, just in case...
		alert(err);
	}
})();