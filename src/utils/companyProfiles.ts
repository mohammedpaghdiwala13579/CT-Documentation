import { CompanyProfile, CompanyId } from "../types";

export const COMPANY_PROFILES: Record<CompanyId, CompanyProfile> = {
  comilla: {
    id: "comilla",
    name: "COMILLA TRADERS",
    tagline1: "Ship Chandler, Marine Supplier & General Merchant",
    tagline2: "Mechanical & Electrical Marine Engineering Services",
    logoUrl: "https://i.ibb.co.com/gFBkpt8B/Chat-GPT-Image-Apr-23-2026-01-10-13-PM.png",
    officeAddress: "Jubilee Road, Chattogram, Bangladesh",
    helplines: "01819315746, 01712-900431",
    email: "comillatraders@gmail.com",
    locationCity: "CHATTOGRAM • BANGLADESH",
    hasStamp: true,
    stampUrl: "https://i.ibb.co.com/jZswrtn6/image-4-removebg-preview.png",
    watermarkUrl: "https://i.ibb.co.com/3mNycQXx/1.png",
    signatureForLabel: "For Comilla Traders",
    firebaseCollection: "documents",
    idPrefix: "doc-",
  },
};
