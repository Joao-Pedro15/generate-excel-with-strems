import ExcelJS from 'exceljs';
import { PassThrough } from 'stream';
import { writeFile } from 'fs/promises'
import { randomInt } from 'crypto';

export async function gerarPlanilhaComoBuffer(data: Array<{ name: string, email: string, id: number }>): Promise<Buffer> {
  const stream = new PassThrough();
  const chunks: Buffer[] = [];

  stream.on('data', (chunk) => chunks.push(chunk));

  const writePromise = new Promise<void>((resolve, reject) => {
    stream.on('end', resolve);
    stream.on('error', reject);
  });

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
  const worksheet = workbook.addWorksheet('Dados');

  // Define colunas
  worksheet.columns = [
    { header: 'ID', key: 'id', width: 10 },
    { header: 'Nome', key: 'name', width: 30 },
    { header: 'Email', key: 'email', width: 30 },
  ];

  // Estiliza o header manualmente (linha 1)
  const headerRow = worksheet.getRow(1);
  headerRow.commit();

  // Adiciona e estiliza dados

  for (const item of data) {
    const row = worksheet.addRow(item);
    row.commit(); // Importante no modo stream
  }

  worksheet.commit(); // Finaliza a worksheet
  await workbook.commit(); // Finaliza o workbook

  await writePromise;
  return Buffer.concat(chunks);
}

function gerarDados(qtd: number) {
  return Array.from({ length: qtd }, (_, i) => {
    return { id: i++, name: `User ${i++}`, email: `User${i++}@email.com` }
  })
}

(async () => {

  const data = gerarDados(1000)
  const buffer = await gerarPlanilhaComoBuffer(data)
  await writeFile('teste.xlsx', buffer)
})()
