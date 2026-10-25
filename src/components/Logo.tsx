import { LOGO_PARTS, LOGO_VIEWBOX } from "./logoPaths";

type Props = { className?: string; animated?: boolean } & React.SVGProps<SVGSVGElement>;

/** Vector LUXEL logo. `animated` adds the parts' intro classes and the light sweep. */
export function Logo({ className, animated, ...rest }: Props) {
  return (
    <svg viewBox={LOGO_VIEWBOX} className={className} style={{ aspectRatio: "1412 / 376" }} role="img" aria-label="LUXEL" {...rest}>
      {animated && (
        <defs>
          <clipPath id="lx-clip">
            {LOGO_PARTS.map((p) => (
              <path key={p.id} d={p.d} />
            ))}
          </clipPath>
          <linearGradient id="lx-sheen-grad" x1="0" x2="1" y1="0" y2="0">
            <stop offset="0" stopColor="#fff" stopOpacity="0" />
            <stop offset="0.5" stopColor="#fff" stopOpacity="0.55" />
            <stop offset="1" stopColor="#fff" stopOpacity="0" />
          </linearGradient>
        </defs>
      )}
      <g fill="#E92620" fillRule="evenodd">
        {LOGO_PARTS.map((p, i) => (
          <path
            key={p.id}
            d={p.d}
            className={animated ? `lx-part lx-${p.id}` : undefined}
            style={animated ? ({ "--i": i } as React.CSSProperties) : undefined}
          />
        ))}
      </g>
      {animated && (
        <g clipPath="url(#lx-clip)">
          <rect className="lx-sheen" x="-360" y="-20" width="320" height="416" fill="url(#lx-sheen-grad)" transform="skewX(-18)" />
        </g>
      )}
    </svg>
  );
}
