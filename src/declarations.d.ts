declare module "*.scss" {
    const content: Record<string, string>;
    export default content;
}

/**
 * siyuan@1.2.4 类型定义未声明 Plugin.getEditor()，
 * 但运行时 API 存在（官方 sample 亦直接使用）。
 * index.ts 通过 currentEditor() 做运行时安全调用，类型在此声明。
 */
interface IPluginEditor {
    protyle: {
        block: { rootID: string };
        notebookId: string;
        path: string;
        [key: string]: any;
    };
    [key: string]: any;
}
