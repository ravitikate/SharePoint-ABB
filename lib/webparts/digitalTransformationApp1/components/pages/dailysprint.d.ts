import "../styles/common.css";
import "../styles/dailySprint.css";
interface IPageModule {
    content: string;
    style?: string;
    script?: () => void;
}
declare const sprintPage: IPageModule;
export default sprintPage;
export { sprintPage };
//# sourceMappingURL=dailysprint.d.ts.map