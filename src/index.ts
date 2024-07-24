import fs from 'fs';
import htmlDocx from 'html-docx-js';
import { Blob } from 'buffer';

interface ConvertHtmlToDocXInput {
    htmlContent: string;
    outputPath: string;
}

export const convertHtmlToDocX = async ({ htmlContent, outputPath }: ConvertHtmlToDocXInput): Promise<void> => {
    try {
        const docxContent = htmlDocx.asBlob(htmlContent) as Blob;

        const arrayBuffer = await docxContent.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        fs.writeFileSync(outputPath, buffer);
        console.log(`DocX file has been saved to ${outputPath}`);
    } catch (error) {
        console.error('Error generating DocX file:', error);
        throw error;
    }
};

// Only run this code if the script is executed directly
if (require.main === module) {
    (async () => {
        await convertHtmlToDocX({
            htmlContent: '<h1>Hello World</h1>',
            outputPath: 'output.docx',
        });
    })();
}
