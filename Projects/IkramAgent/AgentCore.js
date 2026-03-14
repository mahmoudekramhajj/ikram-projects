/**
 * AgentCore.gs - المحرك الرئيسي لوكيل إكرام
 * الإصدار 2.1 - مع إصلاح Loop ومعالجة الأزرار
 */

// ═══════════════════════════════════════════════════════════
// المعالج الرئيسي للرسائل
// ═══════════════════════════════════════════════════════════

function processMessage(platform, odooUserId, userName, messageText) {
  try {
    var userId = String(odooUserId);
    var text = String(messageText).trim();
    var session = getUserSession(userId);
    
    Logger.log('=== processMessage ===');
    Logger.log('User: ' + userId + ', Stage: ' + session.stage + ', Lang: ' + session.lang);
    Logger.log('Message: ' + text);
    
    // ═══════════════════════════════════════════════════════════
    // 1. معالجة الأوامر الخاصة
    // ═══════════════════════════════════════════════════════════
    
    if (text === '/start' || text === 'start' || text === '/menu') {
      // إعادة تعيين الجلسة
      session.lang = null;
      session.stage = 'welcome';
      saveUserSession(userId, session);
      
      return {
        text: WELCOME_MESSAGE,
        buttons: LANGUAGE_BUTTONS
      };
    }
    
    // ═══════════════════════════════════════════════════════════
    // 2. معالجة اختيار اللغة
    // ═══════════════════════════════════════════════════════════
    
    if (!session.lang || session.stage === 'welcome') {
      if (text.indexOf('العربية') !== -1 || text === '🇸🇦 العربية') {
        setUserLanguage(userId, 'ar');
        return showMainMenu('ar');
      }
      if (text.indexOf('English') !== -1 || text === '🇬🇧 English') {
        setUserLanguage(userId, 'en');
        return showMainMenu('en');
      }
      if (text.indexOf('Français') !== -1 || text === '🇫🇷 Français') {
        setUserLanguage(userId, 'fr');
        return showMainMenu('fr');
      }
      if (text.indexOf('أخرى') !== -1 || text.indexOf('Other') !== -1) {
        setUserLanguage(userId, 'other');
        setUserStage(userId, 'other_lang');
        return {
          text: MESSAGES['en'].otherLang,
          buttons: ['🔙 Main Menu']
        };
      }
      
      // إذا لم يختر لغة، أعد رسالة الترحيب
      return {
        text: WELCOME_MESSAGE,
        buttons: LANGUAGE_BUTTONS
      };
    }
    
    var lang = session.lang === 'other' ? 'en' : session.lang;
    var t = MESSAGES[lang] || MESSAGES['en'];
    
    // ═══════════════════════════════════════════════════════════
    // 3. معالجة زر العودة للقائمة الرئيسية
    // ═══════════════════════════════════════════════════════════
    
    if (text.indexOf('القائمة الرئيسية') !== -1 || 
        text.indexOf('Main Menu') !== -1 || 
        text.indexOf('Menu principal') !== -1 ||
        text === '🔙 القائمة الرئيسية' ||
        text === '🔙 Main Menu' ||
        text === '🔙 Menu principal') {
      setUserStage(userId, 'main_menu');
      return showMainMenu(lang);
    }
    
    // ═══════════════════════════════════════════════════════════
    // 4. معالجة القائمة الرئيسية
    // ═══════════════════════════════════════════════════════════
    
    if (session.stage === 'main_menu' || session.stage === 'welcome') {
      
      // باقات حسب الدولة
      if (text.indexOf('باقات حسب الدولة') !== -1 || 
          text.indexOf('Packages by Country') !== -1 ||
          text.indexOf('Forfaits par pays') !== -1) {
        setUserStage(userId, 'select_country');
        return showCountriesList(lang);
      }
      
      // باقات حسب الفندق
      if (text.indexOf('باقات حسب الفندق') !== -1 || 
          text.indexOf('Packages by Hotel') !== -1 ||
          text.indexOf('Forfaits par hôtel') !== -1) {
        setUserStage(userId, 'select_hotel');
        return showHotelsList(lang);
      }
      
      // باقات حسب مدينة الطيران
      if (text.indexOf('باقات حسب مدينة') !== -1 || 
          text.indexOf('Packages by Flight') !== -1 ||
          text.indexOf('Forfaits par ville') !== -1) {
        setUserStage(userId, 'select_city');
        return showCitiesList(lang);
      }
      
      // تواصل معنا
      if (text.indexOf('تواصل') !== -1 || 
          text.indexOf('Contact') !== -1 ||
          text.indexOf('Contactez') !== -1) {
        return {
          text: t.contactUs,
          buttons: [t.backToMenu]
        };
      }
    }
    
    // ═══════════════════════════════════════════════════════════
    // 5. معالجة اختيار الدولة
    // ═══════════════════════════════════════════════════════════
    
    if (session.stage === 'select_country') {
      var countryName = extractCountryFromText(text);
      if (countryName) {
        var packages = getPackagesByCountry(countryName);
        if (packages.length > 0) {
          setUserStage(userId, 'main_menu');
          return {
            text: formatPackagesList(packages, lang, countryName),
            buttons: [t.backToMenu]
          };
        }
      }
      // لم يجد باقات أو الدولة غير صحيحة
      setUserStage(userId, 'main_menu');
      return {
        text: t.noPackages,
        buttons: [t.backToMenu]
      };
    }
    
    // ═══════════════════════════════════════════════════════════
    // 6. معالجة اختيار الفندق
    // ═══════════════════════════════════════════════════════════
    
    if (session.stage === 'select_hotel') {
      var hotelName = text.replace(/[🕋🕌]/g, '').trim();
      if (hotelName && hotelName !== '') {
        var packages = getPackagesByHotel(hotelName);
        if (packages.length > 0) {
          setUserStage(userId, 'main_menu');
          return {
            text: formatPackagesList(packages, lang, hotelName),
            buttons: [t.backToMenu]
          };
        }
      }
      setUserStage(userId, 'main_menu');
      return {
        text: t.noPackages,
        buttons: [t.backToMenu]
      };
    }
    
    // ═══════════════════════════════════════════════════════════
    // 7. معالجة اختيار المدينة
    // ═══════════════════════════════════════════════════════════
    
    if (session.stage === 'select_city') {
      var cityName = text.replace(/✈️/g, '').trim();
      if (cityName && cityName !== '') {
        var packages = getPackagesByCity(cityName);
        if (packages.length > 0) {
          setUserStage(userId, 'main_menu');
          return {
            text: formatPackagesList(packages, lang, cityName),
            buttons: [t.backToMenu]
          };
        }
      }
      setUserStage(userId, 'main_menu');
      return {
        text: t.noPackages,
        buttons: [t.backToMenu]
      };
    }
    
    // ═══════════════════════════════════════════════════════════
    // 8. معالجة لغة أخرى (AI)
    // ═══════════════════════════════════════════════════════════
    
    if (session.stage === 'other_lang') {
      return handleWithAI(text, userId);
    }
    
    // ═══════════════════════════════════════════════════════════
    // 9. الحالة الافتراضية - عرض القائمة
    // ═══════════════════════════════════════════════════════════
    
    return showMainMenu(lang);
    
  } catch (e) {
    Logger.log('processMessage Error: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return {
      text: 'عذراً، حدث خطأ. يرجى المحاولة مرة أخرى.\nSorry, an error occurred. Please try again.',
      buttons: ['🔙 القائمة الرئيسية', '🔙 Main Menu']
    };
  }
}

// ═══════════════════════════════════════════════════════════
// عرض القائمة الرئيسية
// ═══════════════════════════════════════════════════════════

function showMainMenu(lang) {
  var t = MESSAGES[lang] || MESSAGES['en'];
  var buttons = MENU_BUTTONS[lang] || MENU_BUTTONS['en'];
  
  return {
    text: t.mainMenu,
    buttons: buttons
  };
}

// ═══════════════════════════════════════════════════════════
// عرض قائمة الدول
// ═══════════════════════════════════════════════════════════

function showCountriesList(lang) {
  var t = MESSAGES[lang] || MESSAGES['en'];
  var countries = getAvailableCountries();
  
  if (countries.length === 0) {
    return {
      text: t.noPackages,
      buttons: [t.backToMenu]
    };
  }
  
  // أخذ أول 10 دول فقط
  var displayCountries = countries.slice(0, 10);
  displayCountries.push(t.backToMenu);
  
  return {
    text: t.chooseCountry,
    buttons: displayCountries
  };
}

// ═══════════════════════════════════════════════════════════
// عرض قائمة الفنادق
// ═══════════════════════════════════════════════════════════

function showHotelsList(lang) {
  var t = MESSAGES[lang] || MESSAGES['en'];
  var hotels = getAvailableHotels();
  
  if (hotels.length === 0) {
    return {
      text: t.noPackages,
      buttons: [t.backToMenu]
    };
  }
  
  var displayHotels = hotels.slice(0, 10);
  displayHotels.push(t.backToMenu);
  
  return {
    text: t.chooseHotel,
    buttons: displayHotels
  };
}

// ═══════════════════════════════════════════════════════════
// عرض قائمة المدن
// ═══════════════════════════════════════════════════════════

function showCitiesList(lang) {
  var t = MESSAGES[lang] || MESSAGES['en'];
  var cities = getAvailableCities();
  
  if (cities.length === 0) {
    return {
      text: t.noPackages,
      buttons: [t.backToMenu]
    };
  }
  
  var displayCities = cities.slice(0, 10);
  displayCities.push(t.backToMenu);
  
  return {
    text: t.chooseCity,
    buttons: displayCities
  };
}

// ═══════════════════════════════════════════════════════════
// تنسيق قائمة الباقات
// ═══════════════════════════════════════════════════════════

function formatPackagesList(packages, lang, filterName) {
  var t = MESSAGES[lang] || MESSAGES['en'];
  
  var header = '📦 *' + filterName + '*\n';
  header += '━━━━━━━━━━━━━━━\n\n';
  
  var body = '';
  var displayPackages = packages.slice(0, 5); // أول 5 باقات فقط
  
  for (var i = 0; i < displayPackages.length; i++) {
    var pkg = displayPackages[i];
    var name = lang === 'ar' ? pkg.name : (pkg.nameEn || pkg.name);
    var hotels = '';
    
    if (pkg.hotel1Name || pkg.hotel1NameEn) {
      hotels = (lang === 'ar' ? pkg.hotel1Name : pkg.hotel1NameEn) || pkg.hotel1Name;
    }
    if (pkg.hotel2Name || pkg.hotel2NameEn) {
      var h2 = (lang === 'ar' ? pkg.hotel2Name : pkg.hotel2NameEn) || pkg.hotel2Name;
      if (h2) hotels += (hotels ? ' + ' : '') + h2;
    }
    
    body += '📦 *' + name + '*\n';
    body += '💰 ' + t.price + ': ' + formatNumber(pkg.price) + ' SAR\n';
    body += '📅 ' + t.duration + ': ' + pkg.noDays + ' ' + t.days + '\n';
    if (hotels) {
      body += '🏨 ' + t.hotels + ': ' + hotels + '\n';
    }
    body += '✅ ' + t.available + ' (' + pkg.remaining + ' ' + t.remaining + ')\n';
    body += '\n';
  }
  
  if (packages.length > 5) {
    body += '...' + (lang === 'ar' ? ' و ' : ' and ') + (packages.length - 5) + (lang === 'ar' ? ' باقات أخرى' : ' more packages') + '\n';
  }
  
  return header + body;
}

// ═══════════════════════════════════════════════════════════
// استخراج اسم الدولة من النص
// ═══════════════════════════════════════════════════════════

function extractCountryFromText(text) {
  var countries = [
    'Italy', 'Australia', 'Canada', 'USA', 'Switzerland', 'Austria',
    'Germany', 'Netherland', 'Netherlands', 'France', 'UK', 'Spain',
    'Denmark', 'Belgium', 'Sweden', 'South Africa', 'Greece', 'Japan',
    'Finland', 'Norway'
  ];
  
  var cleanText = text.replace(/[🇮🇹🇦🇺🇨🇦🇺🇸🇨🇭🇦🇹🇩🇪🇳🇱🇫🇷🇬🇧🇪🇸🇩🇰🇧🇪🇸🇪🇿🇦🇬🇷🇯🇵🇫🇮🇳🇴🌍]/gu, '').trim();
  
  for (var i = 0; i < countries.length; i++) {
    if (cleanText.toLowerCase() === countries[i].toLowerCase()) {
      return countries[i];
    }
    if (cleanText.toLowerCase().indexOf(countries[i].toLowerCase()) !== -1) {
      return countries[i];
    }
  }
  
  return cleanText;
}

// ═══════════════════════════════════════════════════════════
// تنسيق الأرقام
// ═══════════════════════════════════════════════════════════

function formatNumber(num) {
  if (!num) return '0';
  return Number(num).toLocaleString('en-US');
}

// ═══════════════════════════════════════════════════════════
// معالجة اللغات الأخرى بالذكاء الاصطناعي
// ═══════════════════════════════════════════════════════════

function handleWithAI(text, odooUserId) {
  try {
    var settings = getAllAgentSettings();
    var apiKey = settings.CLAUDE_API_KEY;
    
    if (!apiKey) {
      return {
        text: 'AI service is not configured. Please contact support.',
        buttons: ['🔙 Main Menu']
      };
    }
    
    var packagesSummary = getPackagesSummaryForAI();
    
    var systemPrompt = 'You are a helpful assistant for Ikram AlDyf Hajj company. ' +
      'Answer questions about Hajj packages in the same language the user writes. ' +
      'Available packages:\n' + packagesSummary + '\n\n' +
      'Contact: +966125111940, info@ikramhajj.com';
    
    var response = callClaudeAPI(apiKey, systemPrompt, text);
    
    return {
      text: response,
      buttons: ['🔙 Main Menu']
    };
    
  } catch (e) {
    Logger.log('handleWithAI Error: ' + e.toString());
    return {
      text: 'Sorry, AI service is temporarily unavailable.',
      buttons: ['🔙 Main Menu']
    };
  }
}

// ═══════════════════════════════════════════════════════════
// استدعاء Claude API
// ═══════════════════════════════════════════════════════════

function callClaudeAPI(apiKey, systemPrompt, userMessage) {
  var url = 'https://api.anthropic.com/v1/messages';
  
  var payload = {
    model: 'claude-3-haiku-20240307',
    max_tokens: 500,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userMessage }
    ]
  };
  
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.content && result.content[0]) {
    return result.content[0].text;
  }
  
  return 'Sorry, I could not process your request.';
}

// ═══════════════════════════════════════════════════════════
// دالة اختبار
// ═══════════════════════════════════════════════════════════

function testAgentCore() {
  Logger.log('=== اختبار AgentCore ===');
  
  // اختبار اختيار اللغة
  var result1 = processMessage('Test', 'test123', 'Test User', '/start');
  Logger.log('1. /start:');
  Logger.log('   Text: ' + result1.text.substring(0, 50) + '...');
  Logger.log('   Buttons: ' + result1.buttons.join(', '));
  
  // اختبار اختيار العربية
  var result2 = processMessage('Test', 'test123', 'Test User', '🇸🇦 العربية');
  Logger.log('2. اختيار العربية:');
  Logger.log('   Text: ' + result2.text);
  Logger.log('   Buttons: ' + result2.buttons.join(', '));
  
  // اختبار باقات حسب الدولة
  var result3 = processMessage('Test', 'test123', 'Test User', '🌍 باقات حسب الدولة');
  Logger.log('3. باقات حسب الدولة:');
  Logger.log('   Text: ' + result3.text);
  Logger.log('   Buttons: ' + result3.buttons.slice(0, 3).join(', ') + '...');
  
  // اختبار اختيار Italy
  var result4 = processMessage('Test', 'test123', 'Test User', '🇮🇹 Italy');
  Logger.log('4. اختيار Italy:');
  Logger.log('   Text: ' + result4.text.substring(0, 100) + '...');
}