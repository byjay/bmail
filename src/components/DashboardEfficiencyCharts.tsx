import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as d3 from 'd3';
import {
  Activity,
  ArrowDownRight,
  ArrowUpRight,
  Briefcase,
  Clock,
  Database,
  Gauge,
  Layers,
  TrendingUp,
} from 'lucide-react';
import type {
  DailyEfficiencyPoint,
  ProjectProgressPoint,
  ProjectResponseComparisonPoint,
  WorkEfficiencyMetrics,
} from '../utils/dashboardAnalytics';

interface DashboardEfficiencyChartsProps {
  efficiency: WorkEfficiencyMetrics;
  onOpenPstModal: () => void;
  onOpenProjects: () => void;
  onOpenBriefing?: () => void;
}

const STATUS_BAR_COLORS: Record<ProjectProgressPoint['status'], string> = {
  완료: '#107C41',
  '진행 중': '#0F6CBD',
  '후속 조치 필요': '#D97706',
};

function formatResponseHours(hours: number | null): string {
  if (hours === null || !Number.isFinite(hours)) return '-';
  if (hours < 1) return `${Math.max(1, Math.round(hours * 60))}분`;
  if (hours < 24) return `${hours.toFixed(1)}시간`;
  return `${(hours / 24).toFixed(1)}일`;
}

/**
 * D3 Horizontal Progress Bar Chart for Project Progress Status
 */
const D3ProjectProgressChart: React.FC<{
  data: ProjectProgressPoint[];
  onHover: (item: ProjectProgressPoint | null) => void;
}> = ({ data, onHover }) => {
  const svgRef = useRef<SVGSVGElement | null>(null);

  useEffect(() => {
    const svgEl = svgRef.current;
    if (!svgEl) return;

    const svg = d3.select(svgEl);
    svg.selectAll('*').remove();

    if (data.length === 0) return;

    const width = 520;
    const rowHeight = 34;
    const margin = { top: 8, right: 54, bottom: 24, left: 108 };
    const innerWidth = width - margin.left - margin.right;
    const innerHeight = Math.max(80, data.length * rowHeight);
    const height = innerHeight + margin.top + margin.bottom;

    svg.attr('viewBox', `0 0 ${width} ${height}`);

    const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);

    const x = d3.scaleLinear().domain([0, 100]).range([0, innerWidth]);
    const y = d3
      .scaleBand<string>()
      .domain(data.map((d) => d.projectCode))
      .range([0, innerHeight])
      .padding(0.32);

    const gridTicks = [0, 25, 50, 75, 100];
    g.selectAll('.grid-line')
      .data(gridTicks)
      .enter()
      .append('line')
      .attr('x1', (t) => x(t))
      .attr('x2', (t) => x(t))
      .attr('y1', 0)
      .attr('y2', innerHeight)
      .attr('stroke', '#EDEBE9')
      .attr('stroke-width', 1);

    g.selectAll('.x-label')
      .data(gridTicks)
      .enter()
      .append('text')
      .attr('x', (t) => x(t))
      .attr('y', innerHeight + 16)
      .attr('text-anchor', 'middle')
      .attr('font-size', 10)
      .attr('fill', '#616161')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((t) => `${t}%`);

    const rows = g
      .selectAll('.project-row')
      .data(data)
      .enter()
      .append('g')
      .attr('class', 'project-row')
      .style('cursor', 'pointer')
      .on('mouseenter', function (_event, d) {
        d3.select(this).select('.bar-track').attr('fill', '#E6E6E6');
        onHover(d);
      })
      .on('mouseleave', function () {
        d3.select(this).select('.bar-track').attr('fill', '#F0F0F0');
        onHover(null);
      });

    rows
      .append('text')
      .attr('x', -10)
      .attr('y', (d) => (y(d.projectCode) ?? 0) + y.bandwidth() / 2 + 4)
      .attr('text-anchor', 'end')
      .attr('font-size', 11)
      .attr('font-weight', 600)
      .attr('fill', '#242424')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((d) =>
        d.projectCode.length > 12 ? `${d.projectCode.slice(0, 11)}…` : d.projectCode
      );

    rows
      .append('rect')
      .attr('class', 'bar-track')
      .attr('x', 0)
      .attr('y', (d) => y(d.projectCode) ?? 0)
      .attr('width', innerWidth)
      .attr('height', y.bandwidth())
      .attr('rx', 4)
      .attr('fill', '#F0F0F0');

    rows
      .append('rect')
      .attr('x', 0)
      .attr('y', (d) => y(d.projectCode) ?? 0)
      .attr('width', (d) => Math.max(4, x(d.progressPercent)))
      .attr('height', y.bandwidth())
      .attr('rx', 4)
      .attr('fill', (d) => STATUS_BAR_COLORS[d.status]);

    rows
      .append('text')
      .attr('x', innerWidth + 8)
      .attr('y', (d) => (y(d.projectCode) ?? 0) + y.bandwidth() / 2 + 4)
      .attr('text-anchor', 'start')
      .attr('font-size', 11)
      .attr('font-weight', 600)
      .attr('fill', '#242424')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((d) => `${d.progressPercent}%`);

    rows
      .append('title')
      .text(
        (d) =>
          `[${d.projectCode}] ${d.projectName} — 진행률 ${d.progressPercent}% (${d.status}) | 수신 ${d.received} · 발신 ${d.sent} · 후속 ${d.pendingFollowUps}`
      );
  }, [data, onHover]);

  return (
    <svg
      ref={svgRef}
      role="img"
      aria-label="D3 프로젝트 진행 현황 가로 막대 차트"
      className="w-full h-auto block"
    />
  );
};

/**
 * D3 Multi-Series Combo Chart:
 * - Series 1 (Bars, Left Axis): Project Mail Distribution (Received + Sent stacked bars)
 * - Series 2 (Line + Nodes, Right Axis): Mail Processing Response Speed / Responsiveness Index (0-100)
 */
const D3ProjectResponseMultiChart: React.FC<{
  data: ProjectResponseComparisonPoint[];
  onHover: (pt: ProjectResponseComparisonPoint | null) => void;
}> = ({ data, onHover }) => {
  const svgRef = useRef<SVGSVGElement | null>(null);

  useEffect(() => {
    const svgEl = svgRef.current;
    if (!svgEl) return;

    const svg = d3.select(svgEl);
    svg.selectAll('*').remove();

    if (data.length === 0) return;

    const width = 960;
    const height = 248;
    const margin = { top: 20, right: 54, bottom: 40, left: 46 };
    const innerWidth = width - margin.left - margin.right;
    const innerHeight = height - margin.top - margin.bottom;

    svg.attr('viewBox', `0 0 ${width} ${height}`);

    const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);

    const x = d3
      .scaleBand<string>()
      .domain(data.map((d) => d.projectCode))
      .range([0, innerWidth])
      .padding(0.36);

    const maxMail = Math.max(4, d3.max(data, (d) => d.mailCount) ?? 4);
    const yLeft = d3.scaleLinear().domain([0, maxMail]).nice().range([innerHeight, 0]);
    const yRight = d3.scaleLinear().domain([0, 100]).range([innerHeight, 0]);

    // Horizontal grid lines synchronized with right axis (0, 25, 50, 75, 100)
    const rightTicks = [0, 25, 50, 75, 100];
    rightTicks.forEach((tick) => {
      g.append('line')
        .attr('x1', 0)
        .attr('x2', innerWidth)
        .attr('y1', yRight(tick))
        .attr('y2', yRight(tick))
        .attr('stroke', '#EDEBE9')
        .attr('stroke-width', 1);

      g.append('text')
        .attr('x', innerWidth + 8)
        .attr('y', yRight(tick) + 3.5)
        .attr('text-anchor', 'start')
        .attr('font-size', 10)
        .attr('fill', '#0B7A75')
        .attr('font-family', 'var(--fluent-font-mono)')
        .text(`${tick}`);
    });

    // Left Y-axis labels (Mail Volume)
    const leftTicks = yLeft.ticks(4).filter((v) => Number.isInteger(v));
    leftTicks.forEach((tick) => {
      g.append('text')
        .attr('x', -8)
        .attr('y', yLeft(tick) + 3.5)
        .attr('text-anchor', 'end')
        .attr('font-size', 10)
        .attr('fill', '#616161')
        .attr('font-family', 'var(--fluent-font-mono)')
        .text(`${tick}건`);
    });

    // Series 1: Stacked Bars (Received = Brand Blue, Sent = Soft Brand Tint Blue)
    const barWidth = Math.min(44, x.bandwidth());
    const barOffset = (x.bandwidth() - barWidth) / 2;

    const groups = g
      .selectAll('.multi-col')
      .data(data)
      .enter()
      .append('g')
      .attr('class', 'multi-col')
      .style('cursor', 'pointer')
      .on('mouseenter', function (_event, d) {
        d3.select(this).select('.col-highlight').attr('opacity', 1);
        onHover(d);
      })
      .on('mouseleave', function () {
        d3.select(this).select('.col-highlight').attr('opacity', 0);
        onHover(null);
      });

    // Subtle column hover backdrop
    groups
      .append('rect')
      .attr('class', 'col-highlight')
      .attr('x', (d) => (x(d.projectCode) ?? 0) - 4)
      .attr('y', 0)
      .attr('width', x.bandwidth() + 8)
      .attr('height', innerHeight)
      .attr('rx', 4)
      .attr('fill', '#EBF3FC')
      .attr('opacity', 0);

    // Received segment (bottom)
    groups
      .append('rect')
      .attr('x', (d) => (x(d.projectCode) ?? 0) + barOffset)
      .attr('y', (d) => yLeft(d.received))
      .attr('width', barWidth)
      .attr('height', (d) => Math.max(0, innerHeight - yLeft(d.received)))
      .attr('rx', 3)
      .attr('fill', '#0F6CBD');

    // Sent segment (stacked on top of received)
    groups
      .append('rect')
      .attr('x', (d) => (x(d.projectCode) ?? 0) + barOffset)
      .attr('y', (d) => yLeft(d.mailCount))
      .attr('width', barWidth)
      .attr('height', (d) => Math.max(0, yLeft(d.received) - yLeft(d.mailCount)))
      .attr('rx', 3)
      .attr('fill', '#62ABF5');

    // X-axis Project Labels + Share %
    groups
      .append('text')
      .attr('x', (d) => (x(d.projectCode) ?? 0) + x.bandwidth() / 2)
      .attr('y', innerHeight + 17)
      .attr('text-anchor', 'middle')
      .attr('font-size', 11)
      .attr('font-weight', 600)
      .attr('fill', '#242424')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((d) =>
        d.projectCode.length > 10 ? `${d.projectCode.slice(0, 9)}…` : d.projectCode
      );

    groups
      .append('text')
      .attr('x', (d) => (x(d.projectCode) ?? 0) + x.bandwidth() / 2)
      .attr('y', innerHeight + 31)
      .attr('text-anchor', 'middle')
      .attr('font-size', 9.5)
      .attr('fill', '#616161')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((d) => `${d.mailCount}건 (${d.sharePercent}%)`);

    // Series 2: Response Speed Score Line Overlay (Right Axis)
    const speedLine = d3
      .line<ProjectResponseComparisonPoint>()
      .x((d) => (x(d.projectCode) ?? 0) + x.bandwidth() / 2)
      .y((d) => yRight(d.responseSpeedScore))
      .curve(d3.curveMonotoneX);

    if (data.length > 1) {
      g.append('path')
        .datum(data)
        .attr('d', speedLine)
        .attr('fill', 'none')
        .attr('stroke', '#0B7A75')
        .attr('stroke-width', 2.25);
    }

    // Series 2 Nodes + Speed Value Label
    const nodes = g
      .selectAll('.speed-node')
      .data(data)
      .enter()
      .append('g')
      .attr('class', 'speed-node')
      .style('pointer-events', 'none');

    nodes
      .append('circle')
      .attr('cx', (d) => (x(d.projectCode) ?? 0) + x.bandwidth() / 2)
      .attr('cy', (d) => yRight(d.responseSpeedScore))
      .attr('r', 4.5)
      .attr('fill', '#FFFFFF')
      .attr('stroke', '#0B7A75')
      .attr('stroke-width', 2.25);

    nodes
      .append('text')
      .attr('x', (d) => (x(d.projectCode) ?? 0) + x.bandwidth() / 2)
      .attr('y', (d) => Math.max(10, yRight(d.responseSpeedScore) - 9))
      .attr('text-anchor', 'middle')
      .attr('font-size', 9.5)
      .attr('font-weight', 600)
      .attr('fill', '#0B7A75')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((d) =>
        d.avgResponseHours !== null
          ? formatResponseHours(d.avgResponseHours)
          : `${d.responseSpeedScore}점`
      );

    groups
      .append('title')
      .text(
        (d) =>
          `[${d.projectCode}] ${d.projectName} | 메일 분포 ${d.mailCount}건 (${d.sharePercent}%, 수신 ${d.received}·발신 ${d.sent}) | 반응 속도 지수 ${d.responseSpeedScore}점${
            d.avgResponseHours !== null
              ? ` (평균 회신 ${formatResponseHours(d.avgResponseHours)})`
              : ''
          }`
      );
  }, [data, onHover]);

  return (
    <svg
      ref={svgRef}
      role="img"
      aria-label="D3 프로젝트별 메일 분포 및 메일 처리 반응 속도 다중 시리즈 차트"
      className="w-full h-auto block"
    />
  );
};

/**
 * D3 Radial Arc Gauge for Overall Mail Processing Efficiency
 */
const D3EfficiencyGauge: React.FC<{ percent: number }> = ({ percent }) => {
  const svgRef = useRef<SVGSVGElement | null>(null);

  useEffect(() => {
    const svgEl = svgRef.current;
    if (!svgEl) return;

    const svg = d3.select(svgEl);
    svg.selectAll('*').remove();

    const size = 144;
    const radius = 58;
    const thickness = 13;
    const startAngle = -Math.PI * 0.78;
    const endAngle = Math.PI * 0.78;
    const clamped = Math.max(0, Math.min(100, percent));
    const valueAngle = startAngle + ((endAngle - startAngle) * clamped) / 100;

    svg.attr('viewBox', `0 0 ${size} ${size}`);

    const g = svg.append('g').attr('transform', `translate(${size / 2},${size / 2 + 6})`);

    const trackArc = d3
      .arc<unknown>()
      .innerRadius(radius - thickness)
      .outerRadius(radius)
      .startAngle(startAngle)
      .endAngle(endAngle)
      .cornerRadius(6);

    const fillArc = d3
      .arc<unknown>()
      .innerRadius(radius - thickness)
      .outerRadius(radius)
      .startAngle(startAngle)
      .endAngle(valueAngle)
      .cornerRadius(6);

    g.append('path').attr('d', trackArc({}) || '').attr('fill', '#EDEBE9');

    const gaugeColor =
      clamped >= 75 ? '#107C41' : clamped >= 45 ? '#0F6CBD' : clamped > 0 ? '#D97706' : '#BDBDBD';

    if (clamped > 0) {
      g.append('path').attr('d', fillArc({}) || '').attr('fill', gaugeColor);
    }

    g.append('text')
      .attr('y', 4)
      .attr('text-anchor', 'middle')
      .attr('font-size', 24)
      .attr('font-weight', 700)
      .attr('fill', '#242424')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text(`${clamped}%`);

    g.append('text')
      .attr('y', 22)
      .attr('text-anchor', 'middle')
      .attr('font-size', 10)
      .attr('font-weight', 500)
      .attr('fill', '#616161')
      .text('종합 처리 효율');
  }, [percent]);

  return (
    <svg
      ref={svgRef}
      role="img"
      aria-label={`종합 메일 처리 효율 ${percent}% 게이지`}
      className="w-36 h-36 shrink-0"
    />
  );
};

/**
 * D3 Area + Line Chart for Daily Mail Processing Efficiency Trend
 */
const D3DailyEfficiencyTrend: React.FC<{
  data: DailyEfficiencyPoint[];
  onHover: (pt: DailyEfficiencyPoint | null) => void;
}> = ({ data, onHover }) => {
  const svgRef = useRef<SVGSVGElement | null>(null);

  useEffect(() => {
    const svgEl = svgRef.current;
    if (!svgEl) return;

    const svg = d3.select(svgEl);
    svg.selectAll('*').remove();

    if (data.length === 0) return;

    const width = 440;
    const height = 148;
    const margin = { top: 14, right: 16, bottom: 26, left: 36 };
    const innerWidth = width - margin.left - margin.right;
    const innerHeight = height - margin.top - margin.bottom;

    svg.attr('viewBox', `0 0 ${width} ${height}`);

    const defs = svg.append('defs');
    const grad = defs
      .append('linearGradient')
      .attr('id', 'fluent-d3-eff-grad')
      .attr('x1', '0%')
      .attr('y1', '0%')
      .attr('x2', '0%')
      .attr('y2', '100%');
    grad.append('stop').attr('offset', '0%').attr('stop-color', '#0F6CBD').attr('stop-opacity', 0.22);
    grad.append('stop').attr('offset', '100%').attr('stop-color', '#0F6CBD').attr('stop-opacity', 0.02);

    const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);

    const x = d3
      .scalePoint<string>()
      .domain(data.map((d) => d.shortLabel))
      .range([0, innerWidth])
      .padding(0.25);

    const y = d3.scaleLinear().domain([0, 100]).range([innerHeight, 0]);

    [0, 50, 100].forEach((tick) => {
      g.append('line')
        .attr('x1', 0)
        .attr('x2', innerWidth)
        .attr('y1', y(tick))
        .attr('y2', y(tick))
        .attr('stroke', '#EDEBE9')
        .attr('stroke-width', 1);

      g.append('text')
        .attr('x', -6)
        .attr('y', y(tick) + 3)
        .attr('text-anchor', 'end')
        .attr('font-size', 9.5)
        .attr('fill', '#616161')
        .attr('font-family', 'var(--fluent-font-mono)')
        .text(`${tick}%`);
    });

    const area = d3
      .area<DailyEfficiencyPoint>()
      .x((d) => x(d.shortLabel) ?? 0)
      .y0(innerHeight)
      .y1((d) => y(d.efficiencyPercent))
      .curve(d3.curveMonotoneX);

    const line = d3
      .line<DailyEfficiencyPoint>()
      .x((d) => x(d.shortLabel) ?? 0)
      .y((d) => y(d.efficiencyPercent))
      .curve(d3.curveMonotoneX);

    if (data.length > 1) {
      g.append('path').datum(data).attr('d', area).attr('fill', 'url(#fluent-d3-eff-grad)');
      g.append('path')
        .datum(data)
        .attr('d', line)
        .attr('fill', 'none')
        .attr('stroke', '#0F6CBD')
        .attr('stroke-width', 2);
    }

    const points = g
      .selectAll('.eff-pt')
      .data(data)
      .enter()
      .append('g')
      .attr('class', 'eff-pt')
      .style('cursor', 'pointer')
      .on('mouseenter', function (_event, d) {
        d3.select(this).select('.pt-dot').attr('r', 5.5);
        onHover(d);
      })
      .on('mouseleave', function () {
        d3.select(this).select('.pt-dot').attr('r', 3.5);
        onHover(null);
      });

    points
      .append('circle')
      .attr('class', 'pt-dot')
      .attr('cx', (d) => x(d.shortLabel) ?? 0)
      .attr('cy', (d) => y(d.efficiencyPercent))
      .attr('r', 3.5)
      .attr('fill', '#FFFFFF')
      .attr('stroke', '#0F6CBD')
      .attr('stroke-width', 2);

    points
      .append('circle')
      .attr('cx', (d) => x(d.shortLabel) ?? 0)
      .attr('cy', (d) => y(d.efficiencyPercent))
      .attr('r', 12)
      .attr('fill', 'transparent');

    points
      .append('text')
      .attr('x', (d) => x(d.shortLabel) ?? 0)
      .attr('y', innerHeight + 17)
      .attr('text-anchor', 'middle')
      .attr('font-size', 9.5)
      .attr('fill', '#616161')
      .attr('font-family', 'var(--fluent-font-mono)')
      .text((d) => d.shortLabel);

    points
      .append('title')
      .text(
        (d) =>
          `${d.date}: 처리 효율 ${d.efficiencyPercent}% | 완료 업무 ${d.completedWork}건 · 후속 대기 ${d.pendingFollowUps}건`
      );
  }, [data, onHover]);

  return (
    <svg
      ref={svgRef}
      role="img"
      aria-label="D3 일자별 메일 처리 효율 추이 차트"
      className="w-full h-auto block"
    />
  );
};

export const DashboardEfficiencyCharts: React.FC<DashboardEfficiencyChartsProps> = ({
  efficiency,
  onOpenPstModal,
  onOpenProjects,
  onOpenBriefing,
}) => {
  const [hoveredProject, setHoveredProject] = useState<ProjectProgressPoint | null>(null);
  const [hoveredDay, setHoveredDay] = useState<DailyEfficiencyPoint | null>(null);
  const [hoveredMulti, setHoveredMulti] = useState<ProjectResponseComparisonPoint | null>(null);

  const { monthlyTrend, projectResponseComparison } = efficiency;

  const subMetrics = useMemo(
    () => [
      {
        label: '수신 메일 읽음률',
        value: efficiency.readRatePercent,
        detail: `${efficiency.readInboxCount}/${efficiency.totalInboxCount}건`,
        color: 'var(--fluent-brand)',
      },
      {
        label: '발신 · 회신 대응률',
        value: efficiency.responseRatePercent,
        detail: `발신 ${efficiency.sentCount}건`,
        color: 'var(--fluent-teal)',
      },
      {
        label: '일간 업무 소화율',
        value: efficiency.followUpClosurePercent,
        detail: `완료 ${efficiency.completedWorkCount} · 대기 ${efficiency.pendingFollowUpCount}`,
        color: 'var(--fluent-success)',
      },
    ],
    [efficiency]
  );

  return (
    <div className="space-y-5">
      {/* Row A: Monthly Mail Processing Trend Summary KPI Cards */}
      <section aria-label="월간 메일 처리 트렌드 요약 지표" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
        <div className="fluent-card p-4 flex flex-col justify-between gap-2">
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-medium text-[var(--fluent-fg-3)]">
              월간 메일 처리량 ({monthlyTrend.currentMonthLabel})
            </span>
            <Activity className="w-4 h-4 text-[var(--fluent-brand)] shrink-0" strokeWidth={1.75} />
          </div>
          <div className="flex items-baseline justify-between gap-2">
            <span className="text-2xl font-bold text-[var(--fluent-fg-1)] font-mono tabular-nums">
              {monthlyTrend.currentMonthTotal}
              <span className="text-xs font-normal text-[var(--fluent-fg-3)] ml-0.5">건</span>
            </span>
            {monthlyTrend.momDeltaPercent !== null ? (
              <span
                className={`inline-flex items-center gap-0.5 text-[11px] font-semibold font-mono tabular-nums px-1.5 py-0.5 rounded-[4px] ${
                  monthlyTrend.momDeltaPercent >= 0
                    ? 'bg-[var(--fluent-brand-tint)] text-[var(--fluent-brand)]'
                    : 'bg-[var(--fluent-warning-bg)] text-[var(--fluent-warning)]'
                }`}
              >
                {monthlyTrend.momDeltaPercent >= 0 ? (
                  <ArrowUpRight className="w-3 h-3" />
                ) : (
                  <ArrowDownRight className="w-3 h-3" />
                )}
                전월 대비 {monthlyTrend.momDeltaPercent >= 0 ? '+' : ''}
                {monthlyTrend.momDeltaPercent}%
              </span>
            ) : (
              <span className="text-[11px] text-[var(--fluent-fg-3)]">전월 비교 대기</span>
            )}
          </div>
          <div className="text-[11px] text-[var(--fluent-fg-3)] font-mono tabular-nums">
            수신 {monthlyTrend.currentMonthReceived} · 발신 {monthlyTrend.currentMonthSent}
          </div>
        </div>

        <div className="fluent-card p-4 flex flex-col justify-between gap-2">
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-medium text-[var(--fluent-fg-3)]">
              평균 메일 처리 반응 속도
            </span>
            <Clock className="w-4 h-4 text-[var(--fluent-teal)] shrink-0" strokeWidth={1.75} />
          </div>
          <div className="flex items-baseline justify-between gap-2">
            <span className="text-2xl font-bold text-[var(--fluent-teal)] font-mono tabular-nums">
              {formatResponseHours(monthlyTrend.avgResponseHours)}
            </span>
            <span className="text-[11px] font-mono tabular-nums text-[var(--fluent-fg-3)]">
              {monthlyTrend.measuredReplyPairs > 0
                ? `스레드 ${monthlyTrend.measuredReplyPairs}쌍`
                : '회신 쌍 없음'}
            </span>
          </div>
          <div className="text-[11px] text-[var(--fluent-fg-3)] font-mono tabular-nums">
            {monthlyTrend.fastestResponseHours !== null
              ? `최단 회신 ${formatResponseHours(monthlyTrend.fastestResponseHours)}`
              : '수신·발신 스레드 연결 시 자동 측정'}
          </div>
        </div>

        <div className="fluent-card p-4 flex flex-col justify-between gap-2">
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-medium text-[var(--fluent-fg-3)]">
              월간 처리 효율 · 읽음률
            </span>
            <Gauge className="w-4 h-4 text-[var(--fluent-success)] shrink-0" strokeWidth={1.75} />
          </div>
          <div className="flex items-baseline justify-between gap-2">
            <span className="text-2xl font-bold text-[var(--fluent-success)] font-mono tabular-nums">
              {monthlyTrend.monthlyEfficiencyPercent}%
            </span>
            <span className="text-[11px] font-mono tabular-nums text-[var(--fluent-fg-2)]">
              읽음률 {monthlyTrend.monthlyReadRatePercent}%
            </span>
          </div>
          <div className="w-full h-1.5 bg-[var(--fluent-bg-rail)] rounded-full overflow-hidden">
            <div
              className="h-full rounded-full transition-all duration-200"
              style={{
                width: `${monthlyTrend.monthlyEfficiencyPercent}%`,
                backgroundColor: 'var(--fluent-success)',
              }}
            />
          </div>
        </div>

        <div className="fluent-card p-4 flex flex-col justify-between gap-2">
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-medium text-[var(--fluent-fg-3)]">
              최다 집중 프로젝트 비중
            </span>
            <Layers className="w-4 h-4 text-[var(--fluent-brand)] shrink-0" strokeWidth={1.75} />
          </div>
          <div className="flex items-baseline justify-between gap-2">
            <span className="text-2xl font-bold text-[var(--fluent-brand)] font-mono tabular-nums truncate">
              {monthlyTrend.topProjectCode ? `[${monthlyTrend.topProjectCode}]` : '-'}
            </span>
            <span className="text-xs font-semibold font-mono tabular-nums text-[var(--fluent-fg-2)] shrink-0">
              {monthlyTrend.topProjectSharePercent}%
            </span>
          </div>
          <div className="text-[11px] text-[var(--fluent-fg-3)] font-mono tabular-nums">
            활성 누적 {monthlyTrend.monthlySeries.length}개월 추이 집계
          </div>
        </div>
      </section>

      {/* Row B: Multi-Series D3 Chart (Response Speed vs Project Mail Distribution) */}
      <section
        aria-label="D3 메일 처리 반응 속도 및 프로젝트별 메일 분포 다중 시리즈 차트"
        className="fluent-card p-5 space-y-4"
      >
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="text-sm font-semibold text-[var(--fluent-fg-1)] flex items-center gap-1.5">
              <Activity className="w-4 h-4 text-[var(--fluent-brand)]" strokeWidth={1.75} />
              프로젝트별 메일 분포 · 메일 처리 반응 속도 비교 (D3 다중 시리즈)
            </h2>
            <p className="text-xs text-[var(--fluent-fg-3)] mt-0.5">
              좌축 막대(수신·발신 메일 분포)와 우축 라인(스레드 회신 기반 반응 속도 지수 0~100)을 동시에 비교합니다.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3 text-[11px] text-[var(--fluent-fg-3)]">
            <span className="inline-flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-[2px] bg-[#0F6CBD]" />
              수신 분포
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-[2px] bg-[#62ABF5]" />
              발신 분포
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="w-3 h-0.5 bg-[#0B7A75] inline-block" />
              반응 속도 지수
            </span>
          </div>
        </div>

        {projectResponseComparison.length === 0 ? (
          <div className="h-44 flex flex-col items-center justify-center gap-2.5 text-center border border-dashed border-[var(--fluent-stroke-1)] rounded-[6px] bg-[var(--fluent-bg-subtle)] px-4">
            <p className="text-xs text-[var(--fluent-fg-3)] break-keep">
              실제 메일함(PST · EML · POP3 · Microsoft 365)을 연결하면 프로젝트별 메일 분포와 평균 회신 반응 속도 다중 시리즈 차트가 자동 렌더링됩니다.
            </p>
            <button
              type="button"
              onClick={onOpenPstModal}
              className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-semibold text-[var(--fluent-brand)] bg-white border border-[var(--fluent-stroke-1)] hover:bg-[var(--fluent-brand-tint)] hover:border-[var(--fluent-brand-border)] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
            >
              <Database className="w-3.5 h-3.5" strokeWidth={1.75} />
              원본 메일 불러오기
            </button>
          </div>
        ) : (
          <div className="space-y-2.5">
            <D3ProjectResponseMultiChart
              data={projectResponseComparison}
              onHover={setHoveredMulti}
            />
            <div className="min-h-[28px] px-3 py-1.5 rounded-[4px] bg-[var(--fluent-bg-subtle)] border border-[var(--fluent-stroke-subtle)] flex flex-wrap items-center justify-between gap-2 text-[11px]">
              {hoveredMulti ? (
                <>
                  <span className="font-semibold text-[var(--fluent-fg-1)] truncate">
                    [{hoveredMulti.projectCode}] {hoveredMulti.projectName}
                  </span>
                  <span className="font-mono tabular-nums text-[var(--fluent-fg-2)]">
                    분포 {hoveredMulti.mailCount}건 ({hoveredMulti.sharePercent}% · 수신 {hoveredMulti.received} / 발신 {hoveredMulti.sent}) · 반응 속도{' '}
                    <strong className="text-[#0B7A75]">
                      {hoveredMulti.avgResponseHours !== null
                        ? `${formatResponseHours(hoveredMulti.avgResponseHours)} (${hoveredMulti.responseSpeedScore}점)`
                        : `${hoveredMulti.responseSpeedScore}점`}
                    </strong>
                  </span>
                </>
              ) : (
                <span className="text-[var(--fluent-fg-3)]">
                  각 프로젝트 막대나 반응 속도 포인트에 마우스를 올리면 상세 분포와 평균 회신 소요 시간이 표시됩니다.
                </span>
              )}
            </div>
          </div>
        )}
      </section>

      {/* Row C: Existing D3 Project Progress & Mail Processing Efficiency Pair */}
      <section
        aria-label="D3 프로젝트 진행 현황 및 메일 처리 효율성 시각화"
        className="grid grid-cols-1 lg:grid-cols-12 gap-5"
      >
        {/* Left Card: D3 Project Progress Horizontal Bar Chart */}
        <div className="lg:col-span-6 fluent-card p-5 flex flex-col justify-between space-y-4">
          <div className="flex flex-wrap items-start justify-between gap-2">
            <div>
              <h2 className="text-sm font-semibold text-[var(--fluent-fg-1)] flex items-center gap-1.5">
                <Briefcase className="w-4 h-4 text-[var(--fluent-brand)]" strokeWidth={1.75} />
                프로젝트 진행 현황 (D3)
              </h2>
              <p className="text-xs text-[var(--fluent-fg-3)] mt-0.5">
                프로젝트별 읽음·회신·후속 조치 반영 진행률
              </p>
            </div>

            <div className="flex items-center gap-3 text-[11px] text-[var(--fluent-fg-3)]">
              <span className="inline-flex items-center gap-1">
                <span className="w-2 h-2 rounded-xs bg-[#107C41]" /> 완료
              </span>
              <span className="inline-flex items-center gap-1">
                <span className="w-2 h-2 rounded-xs bg-[#0F6CBD]" /> 진행 중
              </span>
              <span className="inline-flex items-center gap-1">
                <span className="w-2 h-2 rounded-xs bg-[#D97706]" /> 후속 필요
              </span>
            </div>
          </div>

          {efficiency.projectProgress.length === 0 ? (
            <div className="h-44 flex flex-col items-center justify-center gap-2.5 text-center border border-dashed border-[var(--fluent-stroke-1)] rounded-[6px] bg-[var(--fluent-bg-subtle)] px-4">
              <p className="text-xs text-[var(--fluent-fg-3)] break-keep">
                실제 메일이나 PST 파일을 연동하면 프로젝트별 진행 현황 D3 막대 차트가 자동 생성됩니다.
              </p>
              <button
                type="button"
                onClick={onOpenPstModal}
                className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-semibold text-[var(--fluent-brand)] bg-white border border-[var(--fluent-stroke-1)] hover:bg-[var(--fluent-brand-tint)] hover:border-[var(--fluent-brand-border)] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
              >
                <Database className="w-3.5 h-3.5" strokeWidth={1.75} />
                PST · EML 불러오기
              </button>
            </div>
          ) : (
            <div className="space-y-3">
              <D3ProjectProgressChart
                data={efficiency.projectProgress}
                onHover={setHoveredProject}
              />
              <div className="min-h-[28px] px-3 py-1.5 rounded-[4px] bg-[var(--fluent-bg-subtle)] border border-[var(--fluent-stroke-subtle)] flex flex-wrap items-center justify-between gap-2 text-[11px]">
                {hoveredProject ? (
                  <>
                    <span className="font-semibold text-[var(--fluent-fg-1)] truncate">
                      [{hoveredProject.projectCode}] {hoveredProject.projectName}
                    </span>
                    <span className="font-mono tabular-nums text-[var(--fluent-fg-2)]">
                      수신 {hoveredProject.received} · 발신 {hoveredProject.sent} · 후속{' '}
                      {hoveredProject.pendingFollowUps} · {hoveredProject.status}
                    </span>
                  </>
                ) : (
                  <>
                    <span className="text-[var(--fluent-fg-3)]">
                      막대에 마우스를 올리면 프로젝트별 송수신·후속 상세가 표시됩니다.
                    </span>
                    <button
                      type="button"
                      onClick={onOpenProjects}
                      className="font-semibold text-[var(--fluent-brand)] hover:underline whitespace-nowrap"
                    >
                      프로젝트 관리 →
                    </button>
                  </>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Right Card: D3 Mail Processing Efficiency Gauge & Daily Trend */}
        <div className="lg:col-span-6 fluent-card p-5 flex flex-col justify-between space-y-4">
          <div className="flex flex-wrap items-start justify-between gap-2">
            <div>
              <h2 className="text-sm font-semibold text-[var(--fluent-fg-1)] flex items-center gap-1.5">
                <Gauge className="w-4 h-4 text-[var(--fluent-brand)]" strokeWidth={1.75} />
                메일 처리 효율성 · 일간 브리핑 성과 (D3)
              </h2>
              <p className="text-xs text-[var(--fluent-fg-3)] mt-0.5">
                읽음 처리·응답 균형·일간 업무 소화 추이
              </p>
            </div>
            {onOpenBriefing && (
              <button
                type="button"
                onClick={onOpenBriefing}
                className="inline-flex items-center gap-1 text-xs font-semibold text-[var(--fluent-brand)] hover:underline whitespace-nowrap"
              >
                <TrendingUp className="w-3.5 h-3.5" strokeWidth={1.75} />
                일간 브리핑 상세 →
              </button>
            )}
          </div>

          <div className="flex flex-col sm:flex-row items-center gap-4">
            <D3EfficiencyGauge percent={efficiency.overallEfficiencyPercent} />

            <div className="flex-1 w-full space-y-2.5">
              {subMetrics.map((m) => (
                <div key={m.label} className="space-y-1">
                  <div className="flex items-center justify-between gap-2 text-xs">
                    <span className="font-medium text-[var(--fluent-fg-1)]">{m.label}</span>
                    <span className="font-mono tabular-nums text-[var(--fluent-fg-2)] text-[11px]">
                      <strong className="text-[var(--fluent-fg-1)]">{m.value}%</strong> ({m.detail})
                    </span>
                  </div>
                  <div className="w-full h-2 bg-[var(--fluent-bg-rail)] rounded-full overflow-hidden">
                    <div
                      className="h-full rounded-full transition-all duration-200"
                      style={{ width: `${m.value}%`, backgroundColor: m.color }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>

          {efficiency.dailyTrend.length === 0 ? (
            <div className="px-3 py-2.5 rounded-[4px] bg-[var(--fluent-bg-subtle)] border border-dashed border-[var(--fluent-stroke-1)] text-[11px] text-[var(--fluent-fg-3)] text-center">
              활동 일자별 메일 데이터가 연결되면 일간 처리 효율 추이 그래프가 표시됩니다.
            </div>
          ) : (
            <div className="space-y-2 pt-2 border-t border-[var(--fluent-stroke-subtle)]">
              <D3DailyEfficiencyTrend data={efficiency.dailyTrend} onHover={setHoveredDay} />
              <div className="min-h-[26px] px-3 py-1 rounded-[4px] bg-[var(--fluent-bg-subtle)] border border-[var(--fluent-stroke-subtle)] flex items-center justify-between gap-2 text-[11px]">
                {hoveredDay ? (
                  <>
                    <span className="font-mono font-semibold text-[var(--fluent-fg-1)]">
                      {hoveredDay.date}
                    </span>
                    <span className="font-mono tabular-nums text-[var(--fluent-fg-2)]">
                      효율 <strong className="text-[var(--fluent-brand)]">{hoveredDay.efficiencyPercent}%</strong> ·
                      완료 {hoveredDay.completedWork}건 · 후속 대기 {hoveredDay.pendingFollowUps}건
                    </span>
                  </>
                ) : (
                  <span className="text-[var(--fluent-fg-3)]">
                    최근 {efficiency.dailyTrend.length}개 활동 일자 기준 일간 업무 처리 효율 추이
                  </span>
                )}
              </div>
            </div>
          )}
        </div>
      </section>
    </div>
  );
};
