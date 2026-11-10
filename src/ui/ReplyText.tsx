import { Link } from "@fluentui/react-components";
import { selectRange } from "../excel/selectRange";
import { linkifyReply } from "../lib/replyLinks";

interface ReplyTextProps {
  readonly text: string;
}

const jump = (address: string): void => {
  // A bad or stale address (deleted sheet) just does nothing; the text stays readable.
  void selectRange(address).catch(() => undefined);
};

/** Assistant reply text where backticked cell addresses are clickable and select that range in Excel. */
export const ReplyText = ({ text }: ReplyTextProps) => (
  <>
    {linkifyReply(text).map((seg, i) =>
      seg.kind === "text" ? (
        <span key={i}>{seg.text}</span>
      ) : (
        <Link key={i} as="button" inline title={`Select ${seg.address} in Excel`} onClick={() => jump(seg.address)}>
          {seg.text}
        </Link>
      ),
    )}
  </>
);
