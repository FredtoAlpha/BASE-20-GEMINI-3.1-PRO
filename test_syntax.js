const fs = require('fs');
const files = [
  'InterfaceV2.html',
  'InterfaceV2_CoreScript.html',
  'InterfaceV2_InitUpgrade.html',
  'InterfaceV2_GroupsModuleV4_Script.html',
  'InterfaceV2_NewStudentModule.html',
  'InterfaceV2_HeaderControls.html',
  'UIComponents.html',
  'TooltipRegistry.html',
  'OnboardingTour.html'
];
let errors = 0;
files.forEach(f => {
  try {
    const c = fs.readFileSync(f, 'utf-8');
    const scripts = [...c.matchAll(/<script>((?:[\s\S])*?)<\/script>/gi)].map(m => m[1]);
    scripts.forEach((s, i) => {
      try {
        // Supprimer toutes les balises Google Apps Script <? ... ?>
        const cleanScript = s.replace(/<\?[\s\S]*?\?>/g, '');
        const vm = require('vm');
        new vm.Script(cleanScript);
      } catch (e) {
        console.error('Syntax error in', f, 'script block', i, ':', e.message);
        const lines = s.split('\n');
        const errLine = e.stack.match(/evalmachine\.\<anonymous\>\:(\d+)/);
        if (errLine) {
          const lNum = parseInt(errLine[1]);
          console.error(`Line ${lNum}:\n`, lines[lNum - 2]);
          console.error(`> ${lines[lNum - 1]}`);
          console.error(lines[lNum] + '\n');
        }
        errors++;
      }
    });
  } catch (e) {
    console.error('File Error:', f, e.message);
  }
});
process.exit(errors > 0 ? 1 : 0);
