export type RenderOptions = {
    imageDir: string;
    contentWidth: number;
    contentHeight: number;
};

export type InputFile = { kind: 'markdown' | 'docx'; path: string };
