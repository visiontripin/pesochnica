// Упаковка подготовленных потоков VBA в Corel .gms:
//   [18-байтный GMS-заголовок][OLE Compound File (CFBF)]
const fs = require('fs');
const path = require('path');
const CFB = require('cfb');

const BUILD = path.join(__dirname, '..', 'build_streams');
const manifest = JSON.parse(fs.readFileSync(path.join(BUILD, 'manifest.json'), 'utf8'));

function readStream(name) {
  return fs.readFileSync(path.join(BUILD, name));
}

// Заголовок выверен по реальному Corel .gms и GMSViewer:
// "GMS\x01", uint32(10), далее фиксированные 10 байт; OLE начинается с 18-го байта.
const GMS_HEADER = Buffer.from([
  0x47, 0x4d, 0x53, 0x01, 0x0a, 0x00, 0x00, 0x00, 0x02, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
]);

const cfb = CFB.utils.cfb_new(); // корень "Root Entry" + служебный seed-поток

// Потоки уровня корня
for (const name of manifest.root_entries) {
  CFB.utils.cfb_add(cfb, '/' + name, readStream(name));
}
// Потоки внутри storage VBA (промежуточный storage создаётся автоматически)
for (const name of manifest.vba_entries) {
  CFB.utils.cfb_add(cfb, '/VBA/' + name, readStream(name));
}
// seed-заглушка cfb-библиотеки пересоздаётся при каждом cfb_add/init_cfb;
// помечаем её типом 0 (unknown) - rebuild_cfb при сериализации её отбрасывает
const seed = CFB.find(cfb, '/\u0001Sh33tJ5');
if (seed) seed.type = 0;

const oleBuf = CFB.write(cfb, { type: 'buffer' });
const gms = Buffer.concat([GMS_HEADER, oleBuf]);
fs.writeFileSync(manifest.out, gms);
console.log('GMS записан:', manifest.out, gms.length, 'байт (заголовок 18 + OLE', oleBuf.length + ')');
