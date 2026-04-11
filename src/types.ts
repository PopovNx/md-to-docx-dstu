export type GeneratorPaths = {
    imageDir: string;
};

export type InputFile =
    | { kind: 'markdown'; path: string }
    | { kind: 'docx'; path: string };
