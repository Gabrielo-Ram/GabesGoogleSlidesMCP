/**
 *  [IMPORTANT]
 * These types are subject to change depending on the data
 * source. Each CSV, Excel Sheet, or Notion database is organized
 * differently——that's to say, with different headers and properties!
 */
export type SlideType = "Paragraph" | "Bullet";

type ParagraphSlideParam = {
  kind: "Paragraph";
  body: string;
};

type GeneralSummarySlideParam = {
  kind: "GeneralSummary";
  foundedYear: number;
  arr: number;
  industry: string;
  burnRate: number;
  exitStrategy: string;
};

type InvestmentDetails = {
  kind: "InvestmentDetails";
  dealStatus: string;
  fundingStage: string;
  investmentAmount: number;
  investmentDate: string;
};

export type AddSlideParam = {
  companyName: string;
  slideType: GeneralSummarySlideParam | ParagraphSlideParam | InvestmentDetails;
  presentationId?: string;
};

export type CompanyData = {
  companyName: string;
  location: string;
  foundedYear: number;
  arr: number;
  industry: string;
  burnRate: number;
  exitStrategy: string;
  dealStatus: string;
  fundingStage: string;
  investmentAmount: number;
  investmentDate: string;
  keyMetrics: string;
  presentationId?: string;
};
