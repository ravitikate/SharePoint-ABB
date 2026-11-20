import "../styles/common.css";
import "../styles/dailySprint.css";
import "sweetalert2/dist/sweetalert2.min.css";
interface IPageModule {
    content: string;
    style?: string;
    script?: () => void;
}
declare const sprintPage: IPageModule;
export default sprintPage;
export { sprintPage };
//# sourceMappingURL=dailysprint.d.ts.map