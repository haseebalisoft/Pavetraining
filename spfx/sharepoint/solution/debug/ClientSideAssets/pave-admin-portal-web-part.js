(()=>{ var __RUSHSTACK_CURRENT_SCRIPT__ = document.currentScript; define("c1a2b3d4-e5f6-7890-abcd-ef1234567890_1.0.0", ["react","react-dom","@microsoft/sp-core-library","@microsoft/sp-property-pane","@microsoft/sp-webpart-base","PaveAdminPortalWebPartStrings","@microsoft/sp-http"], (__WEBPACK_EXTERNAL_MODULE__650__, __WEBPACK_EXTERNAL_MODULE__729__, __WEBPACK_EXTERNAL_MODULE__878__, __WEBPACK_EXTERNAL_MODULE__723__, __WEBPACK_EXTERNAL_MODULE__134__, __WEBPACK_EXTERNAL_MODULE__580__, __WEBPACK_EXTERNAL_MODULE__272__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 178:
/*!*************************************************!*\
  !*** ./lib/shared/ui/customerPortal.module.css ***!
  \*************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../../node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js */ 726);
// Imports


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(":root{--cp-green:#8dc63f;--cp-green-hover:#7ab536;--cp-green-dark:#5c8228;--cp-green-tint:#f1f8e7;--cp-navy:#2f3540;--cp-navy-deep:#252a33;--cp-charcoal:#4a4a4a;--cp-text:#2b2f36;--cp-muted:#8a8a8a;--cp-page:#f5f6f8;--cp-surface:#fff;--cp-border:#e5e7eb;--cp-warning:#f59e0b;--cp-warning-tint:#fff7ed;--cp-danger:#dc2626;--cp-danger-tint:#fef2f2;--cp-radius:12px;--cp-radius-sm:8px;--cp-shadow:0 1px 2px rgba(47,53,64,.04),0 8px 24px rgba(47,53,64,.06);--cp-font:\"Segoe UI\",system-ui,-apple-system,sans-serif}.root_d30fcc18{background:var(--cp-page);border:1px solid var(--cp-border);border-radius:var(--cp-radius);color:var(--cp-text);font-family:var(--cp-font);min-height:520px;overflow:hidden}.header_d30fcc18{align-items:center;background:var(--cp-surface);border-bottom:1px solid var(--cp-border);display:flex;gap:16px;justify-content:space-between;padding:16px 24px}.headerLeft_d30fcc18{align-items:center;display:flex;flex:1;gap:20px;min-width:0}.logo_d30fcc18{flex-shrink:0;height:48px;object-fit:contain;width:auto}.headerCopy_d30fcc18{min-width:0}.portalTitle_d30fcc18{color:var(--cp-text);font-size:1.35rem;font-weight:700;line-height:1.25;margin:0}.welcome_d30fcc18{color:var(--cp-charcoal);font-size:.95rem;margin:4px 0 0}.welcomeName_d30fcc18{color:var(--cp-green);font-weight:700}.tagline_d30fcc18{color:var(--cp-muted);font-size:.8rem;line-height:1.4;margin:4px 0 0}.accessBadges_d30fcc18{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}.badgeMuted_d30fcc18,.badgeOk_d30fcc18,.badge_d30fcc18{align-items:center;background:var(--cp-page);border:1px solid var(--cp-border);border-radius:6px;color:var(--cp-charcoal);display:inline-flex;font-size:.75rem;font-weight:600;line-height:1.2;padding:4px 10px}.badgeOk_d30fcc18{background:var(--cp-green-tint);border-color:#c5e39a;color:var(--cp-green-dark)}.badgeMuted_d30fcc18{background:#f3f4f6;color:var(--cp-muted)}.userMenu_d30fcc18{align-items:center;background:0 0;border:none;border-radius:var(--cp-radius-sm);color:var(--cp-text);cursor:pointer;display:flex;flex-shrink:0;font:inherit;gap:10px;padding:4px 0 4px 8px}.userMenu_d30fcc18:hover{background:var(--cp-page)}.userName_d30fcc18{font-size:.85rem;font-weight:600}.chevron_d30fcc18{color:var(--cp-muted);font-size:.7rem}.avatar_d30fcc18{background:var(--cp-navy);border-radius:50%;color:#fff;display:inline-flex;font-size:.75rem;font-weight:700;height:36px;width:36px}.avatar_d30fcc18,.menuToggle_d30fcc18{align-items:center;justify-content:center}.menuToggle_d30fcc18{background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius-sm);color:var(--cp-navy);cursor:pointer;display:none;height:40px;width:40px}.navBar_d30fcc18{background:var(--cp-navy);display:flex;flex-wrap:wrap;gap:4px;padding:10px 16px}.navItem_d30fcc18{align-items:center;background:0 0;border:none;border-radius:999px;color:hsla(0,0%,100%,.88);cursor:pointer;display:inline-flex;font:inherit;font-size:.85rem;font-weight:600;gap:8px;padding:8px 14px;transition:background .15s ease,color .15s ease}.navItem_d30fcc18:hover{background:hsla(0,0%,100%,.08);color:#fff}.navItemActive_d30fcc18{background:var(--cp-green);color:#fff}.navItemActive_d30fcc18:hover{background:var(--cp-green-hover);color:#fff}.navIcon_d30fcc18{flex-shrink:0;height:16px;width:16px}.body_d30fcc18{padding:20px 24px 28px}.mobileWelcome_d30fcc18{display:none;font-size:.95rem;margin:0 0 14px}.stats_d30fcc18{display:grid;gap:14px;grid-template-columns:repeat(4,minmax(0,1fr));margin-bottom:18px}.statCard_d30fcc18{align-items:flex-start;background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius);box-shadow:var(--cp-shadow);display:flex;gap:10px;justify-content:space-between;padding:16px 16px 14px;transition:transform .15s ease,box-shadow .15s ease}.statCard_d30fcc18:hover{box-shadow:0 10px 28px rgba(47,53,64,.1);transform:translateY(-1px)}.statLabel_d30fcc18{color:var(--cp-muted);font-size:.68rem;font-weight:700;letter-spacing:.06em;margin:0;text-transform:uppercase}.statValue_d30fcc18{color:var(--cp-text);font-size:1.75rem;font-weight:700;line-height:1;margin:8px 0 0}.statHint_d30fcc18{font-size:.75rem;font-weight:600;margin:8px 0 0}.statHintGood_d30fcc18{color:var(--cp-green-dark)}.statHintWarn_d30fcc18{color:var(--cp-warning)}.statIcon_d30fcc18{align-items:center;border-radius:50%;display:inline-flex;flex-shrink:0;height:40px;justify-content:center;width:40px}.statIconGreen_d30fcc18{background:var(--cp-green-tint);color:var(--cp-green-dark)}.statIconOrange_d30fcc18{background:var(--cp-warning-tint);color:var(--cp-warning)}.statIconRed_d30fcc18{background:var(--cp-danger-tint);color:var(--cp-danger)}.card_d30fcc18{background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius);box-shadow:var(--cp-shadow);padding:16px}.cardHeader_d30fcc18{align-items:center;display:flex;flex-wrap:wrap;gap:12px;justify-content:space-between;margin-bottom:14px}.cardTitle_d30fcc18{font-size:1.05rem;font-weight:700;margin:0}.cardActions_d30fcc18{align-items:center;display:flex;flex-wrap:wrap;gap:8px}.linkBtn_d30fcc18{background:0 0;border:none;color:var(--cp-green-dark);cursor:pointer;font:inherit;font-size:.85rem;font-weight:700;padding:4px 2px}.linkBtn_d30fcc18:hover{color:var(--cp-green);text-decoration:underline}.primaryBtn_d30fcc18{background:var(--cp-green);border:none;border-radius:var(--cp-radius-sm);color:#fff;cursor:pointer;font:inherit;font-size:.82rem;font-weight:700;padding:8px 14px}.primaryBtn_d30fcc18:hover{background:var(--cp-green-hover)}.ghostBtn_d30fcc18{align-items:center;background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius-sm);color:var(--cp-charcoal);cursor:pointer;display:inline-flex;font:inherit;font-size:.82rem;font-weight:600;gap:6px;padding:8px 12px}.ghostBtn_d30fcc18:hover{background:var(--cp-green-tint);border-color:var(--cp-green)}.searchInput_d30fcc18{background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius-sm);color:var(--cp-text);font:inherit;font-size:.85rem;min-width:160px;padding:8px 12px}.searchInput_d30fcc18:focus{outline:2px solid var(--cp-green);outline-offset:1px}.dashMain_d30fcc18{display:grid;gap:16px;grid-template-columns:minmax(0,1.55fr) minmax(0,1fr);margin-bottom:16px}.dashSide_d30fcc18{display:flex;flex-direction:column;gap:16px;min-width:0}.dashBottom_d30fcc18{display:grid;gap:16px;grid-template-columns:1fr 1fr}.tableWrap_d30fcc18{overflow-x:auto}.table_d30fcc18{border-collapse:collapse;font-size:.85rem;width:100%}.table_d30fcc18 th{background:#fafbfc;color:var(--cp-muted);font-size:.72rem;font-weight:700;letter-spacing:.04em;padding:10px 12px;text-align:left;text-transform:uppercase}.table_d30fcc18 td,.table_d30fcc18 th{border-bottom:1px solid var(--cp-border)}.table_d30fcc18 td{padding:12px;vertical-align:middle}.table_d30fcc18 tr:hover td{background:#fafbfc}.candidateCell_d30fcc18{align-items:center;display:flex;font-weight:600;gap:10px}.avatarSm_d30fcc18{background:var(--cp-green-tint);border-radius:50%;color:var(--cp-green-dark);flex-shrink:0;font-size:.7rem;height:32px;justify-content:center;width:32px}.avatarSm_d30fcc18,.pill_d30fcc18{align-items:center;display:inline-flex;font-weight:700}.pill_d30fcc18{border-radius:999px;font-size:.72rem;gap:4px;padding:4px 10px;white-space:nowrap}.pillOk_d30fcc18{background:var(--cp-green-tint);color:var(--cp-green-dark)}.pillWarn_d30fcc18{background:var(--cp-warning-tint);color:#b45309}.pillBad_d30fcc18{background:var(--cp-danger-tint);color:var(--cp-danger)}.tableFooter_d30fcc18{align-items:center;display:flex;flex-wrap:wrap;gap:12px;justify-content:space-between;margin-top:12px;padding-top:4px}.muted_d30fcc18{color:var(--cp-muted);font-size:.8rem;margin:0}.error_d30fcc18{color:var(--cp-danger);font-size:.85rem}.docGrid_d30fcc18{display:grid;gap:10px;grid-template-columns:repeat(2,minmax(0,1fr))}.docTile_d30fcc18{background:#fafbfc;border:1px solid var(--cp-border);border-radius:var(--cp-radius-sm);cursor:pointer;display:flex;flex-direction:column;font:inherit;gap:8px;padding:12px 10px;text-align:left;transition:border-color .15s ease,background .15s ease}.docTile_d30fcc18:hover{background:var(--cp-green-tint);border-color:var(--cp-green)}.docIcon_d30fcc18{height:28px;width:28px}.docLabel_d30fcc18{color:var(--cp-text);font-size:.78rem;font-weight:700;line-height:1.3}.docMeta_d30fcc18{color:var(--cp-muted);font-size:.7rem}.pdfTile_d30fcc18{background:var(--cp-danger-tint);border-color:#fecaca}.pdfTile_d30fcc18:hover{background:#fee2e2;border-color:var(--cp-danger)}.nvqList_d30fcc18{display:flex;flex-direction:column;gap:14px}.nvqRow_d30fcc18{display:flex;flex-direction:column;gap:6px}.nvqTop_d30fcc18{align-items:baseline;display:flex;gap:8px;justify-content:space-between}.nvqName_d30fcc18{font-size:.85rem;font-weight:700;margin:0}.nvqCourse_d30fcc18{color:var(--cp-muted);font-size:.75rem;margin:0}.nvqPct_d30fcc18{color:var(--cp-green-dark);font-size:.75rem;font-weight:700}.progressTrack_d30fcc18{background:#eef0f3;border-radius:999px;height:8px;overflow:hidden}.progressFill_d30fcc18{background:var(--cp-green);border-radius:999px;height:100%}.eventList_d30fcc18{display:flex;flex-direction:column;gap:10px}.eventCard_d30fcc18{align-items:center;background:#fafbfc;border:1px solid var(--cp-border);border-radius:var(--cp-radius-sm);cursor:pointer;display:flex;font:inherit;gap:14px;padding:12px;text-align:left;transition:border-color .15s ease,background .15s ease;width:100%}.eventCard_d30fcc18:hover{background:var(--cp-green-tint);border-color:var(--cp-green)}.dateBadge_d30fcc18{background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius-sm);flex-shrink:0;padding:6px 4px;text-align:center;width:52px}.dateMonth_d30fcc18{color:var(--cp-muted);display:block;font-size:.65rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase}.dateDay_d30fcc18{color:var(--cp-text);display:block;font-size:1.25rem;font-weight:700;line-height:1.1}.eventTitle_d30fcc18{font-size:.9rem;font-weight:700;margin:0}.eventMeta_d30fcc18{color:var(--cp-muted);font-size:.75rem;margin:4px 0 0}.offerGrid_d30fcc18{display:grid;gap:10px;grid-template-columns:1fr 1fr}.offerCard_d30fcc18{background:linear-gradient(135deg,rgba(47,53,64,.92),rgba(37,42,51,.88)),linear-gradient(135deg,#3a4a2a,#2f3540);border:none;border-radius:var(--cp-radius);color:#fff;cursor:pointer;display:flex;flex-direction:column;font:inherit;gap:8px;justify-content:flex-end;min-height:140px;overflow:hidden;padding:16px 14px;position:relative;text-align:left;transition:transform .15s ease}.offerCard_d30fcc18:hover{transform:translateY(-2px)}.offerBadge_d30fcc18{background:var(--cp-green);border-radius:6px;color:#fff;font-size:.72rem;font-weight:800;left:12px;padding:4px 8px;position:absolute;top:12px}.offerTitle_d30fcc18{font-size:.95rem;font-weight:700;line-height:1.3;margin:0}.offerCode_d30fcc18{align-self:flex-start;background:hsla(0,0%,100%,.14);border:1px solid hsla(0,0%,100%,.22);border-radius:999px;display:inline-flex;font-size:.72rem;font-weight:600;padding:4px 10px}.offerLink_d30fcc18{color:var(--cp-green);font-size:.8rem;font-weight:700}.listPage_d30fcc18{display:flex;flex-direction:column;gap:12px}.pageTitle_d30fcc18{font-size:1.25rem;font-weight:700;margin:0}.pageSubtitle_d30fcc18{color:var(--cp-muted);font-size:.875rem;margin:0}.stubBox_d30fcc18{background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:var(--cp-radius);color:var(--cp-muted);padding:20px}.bottomNav_d30fcc18{display:none}@media (max-width:1100px){.dashBottom_d30fcc18,.dashMain_d30fcc18{grid-template-columns:1fr}.stats_d30fcc18{grid-template-columns:repeat(2,minmax(0,1fr))}}@media (max-width:768px){.header_d30fcc18{padding:12px 14px}.headerCopy_d30fcc18,.navBar_d30fcc18,.userMenu_d30fcc18{display:none}.headerLeft_d30fcc18{justify-content:space-between;width:100%}.menuToggle_d30fcc18{display:inline-flex}.mobileWelcome_d30fcc18{display:block}.body_d30fcc18{padding:14px 14px 72px}.stats_d30fcc18{gap:10px;grid-template-columns:repeat(2,minmax(0,1fr))}.offerGrid_d30fcc18{grid-template-columns:1fr}.bottomNav_d30fcc18{background:var(--cp-surface);border-top:1px solid var(--cp-border);bottom:0;display:flex;justify-content:space-around;left:0;margin:0 -14px -14px;padding:6px 4px calc(6px + env(safe-area-inset-bottom,0));position:sticky;right:0;z-index:20}.tabItem_d30fcc18{align-items:center;background:0 0;border:none;color:var(--cp-muted);cursor:pointer;display:flex;flex:1;flex-direction:column;font:inherit;font-size:.65rem;font-weight:600;gap:2px;padding:6px 2px}.tabItemActive_d30fcc18{color:var(--cp-green-dark)}.mobileNavDrawer_d30fcc18{background:var(--cp-navy);display:flex;flex-direction:column;gap:4px;padding:10px 14px 14px}.hideDesktopTableCols_d30fcc18{display:none}}@media (min-width:769px){.mobileNavDrawer_d30fcc18{display:none!important}}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vQzovVXNlcnMvQWJkdWwvRG93bmxvYWRzL05ldyUyMGZvbGRlciUyMCgyKS9zcGZ4L3NyYy9zaGFyZWQvdWkvY3VzdG9tZXJQb3J0YWwubW9kdWxlLnNjc3MiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IkFBR0EsTUFDRSxrQkFBQSxDQUNBLHdCQUFBLENBQ0EsdUJBQUEsQ0FDQSx1QkFBQSxDQUNBLGlCQUFBLENBQ0Esc0JBQUEsQ0FDQSxxQkFBQSxDQUNBLGlCQUFBLENBQ0Esa0JBQUEsQ0FDQSxpQkFBQSxDQUNBLGlCQUFBLENBQ0EsbUJBQUEsQ0FDQSxvQkFBQSxDQUNBLHlCQUFBLENBQ0EsbUJBQUEsQ0FDQSx3QkFBQSxDQUNBLGdCQUFBLENBQ0Esa0JBQUEsQ0FDQSxzRUFBQSxDQUNBLHVEQUFBLENBR0YsZUFHRSx5QkFBQSxDQUlBLGlDQUFBLENBRkEsOEJBQUEsQ0FIQSxvQkFBQSxDQURBLDBCQUFBLENBR0EsZ0JBQUEsQ0FFQSxlQUNBLENBS0YsaUJBRUUsa0JBQUEsQ0FJQSw0QkFBQSxDQUNBLHdDQUFBLENBTkEsWUFBQSxDQUdBLFFBQUEsQ0FEQSw2QkFBQSxDQUVBLGlCQUVBLENBR0YscUJBRUUsa0JBQUEsQ0FEQSxZQUFBLENBSUEsTUFBQSxDQUZBLFFBQUEsQ0FDQSxXQUNBLENBR0YsZUFJRSxhQUFBLENBSEEsV0FBQSxDQUVBLGtCQUFBLENBREEsVUFFQSxDQUdGLHFCQUNFLFdBQUEsQ0FHRixzQkFJRSxvQkFBQSxDQUZBLGlCQUFBLENBQ0EsZUFBQSxDQUVBLGdCQUFBLENBSkEsUUFJQSxDQUdGLGtCQUdFLHdCQUFBLENBREEsZ0JBQUEsQ0FEQSxjQUVBLENBR0Ysc0JBQ0UscUJBQUEsQ0FDQSxlQUFBLENBR0Ysa0JBR0UscUJBQUEsQ0FEQSxlQUFBLENBRUEsZUFBQSxDQUhBLGNBR0EsQ0FHRix1QkFDRSxZQUFBLENBQ0EsY0FBQSxDQUNBLE9BQUEsQ0FDQSxlQUFBLENBR0YsdURBSUUsa0JBQUEsQ0FPQSx5QkFBQSxD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 */", true);


/***/ }),

/***/ 955:
/*!*****************************************!*\
  !*** ./lib/shared/ui/portal.module.css ***!
  \*****************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../../node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js */ 726);
// Imports


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(".shell_c8d82a1a{background:#f7f7f7;border:1px solid #e2e2e2;border-radius:12px;color:#4a4b4d;display:flex;font-family:Segoe UI,system-ui,sans-serif;min-height:520px;overflow:hidden}.adminShell_c8d82a1a{flex-direction:column;max-width:100%;min-width:0;width:100%}.adminTopNav_c8d82a1a{background:#fff;border-bottom:1px solid rgba(74,75,77,.1);position:sticky;top:0;z-index:40}.adminTopNavBar_c8d82a1a{align-items:center;display:flex;gap:12px;min-height:56px;padding:10px 14px}.adminBrandBlock_c8d82a1a{flex-shrink:0;min-width:0}.adminBrandBlock_c8d82a1a .brand_c8d82a1a{font-size:.92rem;letter-spacing:.06em;text-transform:uppercase}.adminBrandBlock_c8d82a1a .tagline_c8d82a1a{font-size:.68rem;margin:2px 0 0}.adminNavDesktop_c8d82a1a{-webkit-overflow-scrolling:touch;align-items:center;display:flex;flex:1;gap:4px;min-width:0;overflow-x:auto;scrollbar-width:thin}.adminNavLink_c8d82a1a{background:0 0;border:0;border-radius:999px;color:#6b6b6b;cursor:pointer;flex-shrink:0;font:inherit;font-size:.82rem;font-weight:650;padding:8px 12px;white-space:nowrap}.adminNavLink_c8d82a1a:hover{background:#eef7e4;color:#4a4b4d}.adminNavLinkActive_c8d82a1a{background:rgba(122,193,67,.22);color:#333436}.adminTopNavTrailing_c8d82a1a{align-items:center;display:flex;flex-shrink:0;gap:8px;margin-left:auto}.adminUserChip_c8d82a1a{background:#f7f7f7;border:1px solid #e2e2e2;border-radius:999px;color:#6b6b6b;display:none;font-size:.75rem;font-weight:600;max-width:180px;overflow:hidden;padding:6px 10px;text-overflow:ellipsis;white-space:nowrap}.adminMenuToggle_c8d82a1a{background:#fff;border:1px solid #e2e2e2;border-radius:999px;color:#4a4b4d;cursor:pointer;display:none;font:inherit;font-size:.8rem;font-weight:700;padding:8px 12px}.adminMobileDrawer_c8d82a1a,.adminMobileScrim_c8d82a1a{display:none}.adminMobileTitle_c8d82a1a{color:#8a8a8a;font-size:.72rem;font-weight:750;letter-spacing:.08em;margin:0 0 8px;text-transform:uppercase}.adminMobileNav_c8d82a1a{display:flex;flex-direction:column;gap:4px}.adminMobileLink_c8d82a1a{background:0 0;border:0;border-radius:10px;color:#4a4b4d;cursor:pointer;font:inherit;font-size:.95rem;font-weight:650;padding:12px;text-align:left}.adminMobileLinkActive_c8d82a1a,.adminMobileLink_c8d82a1a:hover{background:#eef7e4}.sidebar_c8d82a1a{background:#fff;border-right:1px solid #e2e2e2;display:flex;flex-direction:column;flex-shrink:0;gap:4px;padding:16px 10px;width:220px}.brand_c8d82a1a{color:#4a4b4d;font-size:1.15rem;font-weight:800;margin:0}.tagline_c8d82a1a{color:#8a8a8a;font-size:.7rem;font-weight:600;margin:0 0 8px}.chip_c8d82a1a{background:#eef7e4;border-radius:4px;color:#5e9a2f;display:inline-block;font-size:.7rem;font-weight:700;margin-bottom:12px;padding:4px 8px}.navBtn_c8d82a1a{background:0 0;border:none;border-radius:6px;color:#4a4b4d;cursor:pointer;font-size:.85rem;font-weight:600;padding:10px 12px;text-align:left;width:100%}.navBtn_c8d82a1a:hover{background:#eef7e4}.navBtnActive_c8d82a1a{background:#7ac143;color:#fff}.navBtnActive_c8d82a1a:hover{background:#5e9a2f}.main_c8d82a1a{background:#f7f7f7;box-sizing:border-box;flex:1;min-width:0;padding:16px 18px 24px;width:100%}.pageHeader_c8d82a1a{background:#fff;border:1px solid rgba(74,75,77,.08);border-radius:16px;box-shadow:0 8px 22px rgba(42,42,42,.04);margin-bottom:14px;padding:14px 16px}.eyebrow_c8d82a1a{color:#8a8a8a;font-size:.7rem;font-weight:700;letter-spacing:.06em;margin:0 0 4px;text-transform:uppercase}.title_c8d82a1a{color:#333436;font-size:clamp(1.2rem,2.5vw,1.45rem);font-weight:700;letter-spacing:-.02em;margin:0 0 4px}.subtitle_c8d82a1a{color:#8a8a8a;font-size:.875rem;line-height:1.45;margin:0}.stats_c8d82a1a{display:grid;gap:10px;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));margin-bottom:16px}.stat_c8d82a1a{background:linear-gradient(180deg,#fff,#f8f9f6);border:1px solid rgba(74,75,77,.08);border-radius:14px;box-shadow:0 6px 16px rgba(42,42,42,.04);min-width:0;padding:14px}.statValue_c8d82a1a{color:#333436;font-size:clamp(1.2rem,3vw,1.45rem);font-weight:750;letter-spacing:-.02em;margin:0}.statLabel_c8d82a1a{color:#8a8a8a;font-size:.68rem;font-weight:700;letter-spacing:.04em;margin:6px 0 0;overflow-wrap:anywhere;text-transform:uppercase}.panel_c8d82a1a{background:#fff;border:1px solid rgba(74,75,77,.08);border-radius:16px;box-shadow:0 8px 22px rgba(42,42,42,.04);min-width:0;padding:14px}.tableWrap_c8d82a1a{-webkit-overflow-scrolling:touch;border:1px solid rgba(74,75,77,.08);border-radius:10px;max-height:min(70vh,560px);overflow:auto;width:100%}.table_c8d82a1a{border-collapse:collapse;font-size:.85rem;min-width:36rem;width:100%}.table_c8d82a1a td,.table_c8d82a1a th{border-bottom:1px solid #ebebeb;padding:10px 12px;text-align:left;vertical-align:middle}.table_c8d82a1a th{background:#f3f4f2;color:#6b6b6b;font-size:.72rem;font-weight:700;letter-spacing:.04em;position:sticky;text-transform:uppercase;top:0;white-space:nowrap;z-index:2}.table_c8d82a1a td{overflow-wrap:anywhere}.muted_c8d82a1a{color:#8a8a8a;font-size:.875rem}.error_c8d82a1a{color:#dc2626;font-size:.875rem}.toolbar_c8d82a1a{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}.toolbar_c8d82a1a button{background:#fff;border:1px solid #e2e2e2;border-radius:999px;color:#4a4b4d;cursor:pointer;font-size:.82rem;font-weight:650;padding:8px 14px}.toolbar_c8d82a1a button:hover{background:#eef7e4;border-color:#7ac143}.toolbarBtnActive_c8d82a1a{background:#7ac143!important;border-color:#7ac143!important;color:#fff!important}.success_c8d82a1a{color:#5e9a2f;font-size:.875rem;font-weight:600}.actionsCol_c8d82a1a{background:#fff;left:0;position:sticky;white-space:nowrap;z-index:1}.table_c8d82a1a th.actionsCol_c8d82a1a{background:#f3f4f2;z-index:3}.linkBtnDanger_c8d82a1a,.linkBtn_c8d82a1a{background:0 0;border:none;cursor:pointer;font-size:.8rem;font-weight:700;margin-right:4px;padding:2px 6px}.linkBtn_c8d82a1a{color:#5e9a2f}.linkBtnDanger_c8d82a1a{color:#dc2626}.linkBtnDanger_c8d82a1a:disabled,.linkBtn_c8d82a1a:disabled{cursor:not-allowed;opacity:.5}.drawerBackdrop_c8d82a1a{background:rgba(30,30,30,.45);display:flex;inset:0;justify-content:flex-end;position:fixed;z-index:1000}.drawer_c8d82a1a{background:#fff;box-shadow:-4px 0 24px rgba(0,0,0,.12);display:flex;flex-direction:column;height:100%;overflow:auto;padding:16px;width:min(480px,100%)}.drawerHeader_c8d82a1a{align-items:flex-start;display:flex;gap:12px;justify-content:space-between;margin-bottom:8px}.formGrid_c8d82a1a{display:flex;flex:1;flex-direction:column;gap:10px;margin:12px 0}.field_c8d82a1a{display:flex;flex-direction:column;gap:4px}.fieldLabel_c8d82a1a{color:#8a8a8a;font-size:.75rem;font-weight:700}.field_c8d82a1a input,.field_c8d82a1a select{border:1px solid #e2e2e2;border-radius:8px;box-sizing:border-box;color:#4a4b4d;font-size:.875rem;min-height:40px;padding:10px 12px}.field_c8d82a1a input:focus,.field_c8d82a1a select:focus{outline:2px solid rgba(122,193,67,.45);outline-offset:1px}.drawerFooter_c8d82a1a{border-top:1px solid #e2e2e2;display:flex;flex-wrap:wrap;gap:8px;padding-top:12px}.drawerFooter_c8d82a1a button{background:#fff;border:1px solid #e2e2e2;border-radius:999px;color:#4a4b4d;cursor:pointer;font-weight:700;padding:10px 16px}.drawerFooter_c8d82a1a button:first-child{background:#7ac143;border-color:#7ac143;color:#fff}.drawerFooter_c8d82a1a button:disabled{cursor:not-allowed;opacity:.6}.logoBox_c8d82a1a{background:#f8fbf2;border:1px dashed #c5d6a7;border-radius:12px;display:flex;flex-direction:column;gap:8px;margin:12px 0;padding:12px}.logoPreview_c8d82a1a{background:#fff;border:1px solid #e2e2e2;border-radius:8px;height:96px;object-fit:contain;width:96px}.logoBox_c8d82a1a input[type=file]{font-size:.8rem}@media (min-width:960px){.adminUserChip_c8d82a1a{display:inline-block}.adminMenuToggle_c8d82a1a{display:none}}@media (max-width:959px){.adminNavDesktop_c8d82a1a{display:none}.adminMenuToggle_c8d82a1a{display:inline-flex}.adminMobileScrim_c8d82a1a{background:rgba(20,24,18,.42);display:block;inset:0;position:fixed;z-index:50}.adminMobileDrawer_c8d82a1a{background:#fff;bottom:0;box-shadow:-16px 0 40px rgba(26,26,26,.18);display:flex;flex-direction:column;padding:16px;position:fixed;right:0;top:0;transform:translateX(105%);transition:transform .18s ease;visibility:hidden;width:min(20rem,88vw);z-index:60}.adminMobileDrawerOpen_c8d82a1a{transform:translateX(0);visibility:visible}.adminMobileDrawer_c8d82a1a[hidden]{display:none!important}.adminMobileDrawerOpen_c8d82a1a[hidden]{display:flex!important}.main_c8d82a1a{padding:12px 12px 20px}.pageHeader_c8d82a1a{border-radius:14px;padding:12px 14px}.toolbar_c8d82a1a button{flex:1 1 auto;min-width:calc(50% - 8px);text-align:center}.drawer_c8d82a1a{width:100%}}@media (max-width:560px){.adminTopNavBar_c8d82a1a{padding:8px 10px}.stats_c8d82a1a{grid-template-columns:repeat(2,minmax(0,1fr))}.table_c8d82a1a{font-size:.8rem;min-width:28rem}.table_c8d82a1a td,.table_c8d82a1a th{padding:8px 10px}}.hubPage_c8d82a1a{display:flex;flex-direction:column;gap:1.35rem;min-width:0;width:100%}.hubHero_c8d82a1a{background:linear-gradient(180deg,rgba(20,24,18,.28),rgba(20,24,18,.55) 55%,rgba(20,24,18,.72)),radial-gradient(120% 90% at 15% 10%,rgba(122,193,67,.35),transparent 50%),linear-gradient(135deg,#3d4a32,#5a6b45 40%,#2a3324);border-radius:1.15rem;box-shadow:0 14px 36px rgba(42,42,42,.14);box-sizing:border-box;color:#f5f5f5;min-width:0;overflow:clip;padding:1.6rem 1.35rem 1.25rem;position:relative;width:100%}.hubHeroInner_c8d82a1a{display:grid;gap:1.1rem;min-width:0;position:relative;z-index:1}.hubHeroEyebrow_c8d82a1a{color:hsla(0,0%,96%,.72);font-size:.72rem;font-weight:700;letter-spacing:.08em;margin:0;text-transform:uppercase}.hubHeroTitle_c8d82a1a{color:#fff;font-size:clamp(1.55rem,4vw,2.3rem);font-weight:750;letter-spacing:-.03em;line-height:1.12;margin:0}.hubHeroSubtitle_c8d82a1a{color:hsla(0,0%,100%,.88);font-size:clamp(.92rem,2.5vw,1rem);line-height:1.5;margin:0;max-width:36rem}.hubHeroPanel_c8d82a1a{background:#fff;border-radius:1rem;box-shadow:0 16px 36px rgba(0,0,0,.18);color:#4a4b4d;display:grid;gap:1rem 1.25rem;grid-template-columns:minmax(0,1.2fr) minmax(0,.8fr);min-width:0;padding:1.1rem 1.15rem}.hubHeroPanelLabel_c8d82a1a{color:#8a8a8a;font-size:.7rem;font-weight:700;letter-spacing:.05em;margin:0 0 .35rem;text-transform:uppercase}.hubHeroPanelHint_c8d82a1a{color:#6b6b6b;font-size:.9rem;line-height:1.45;margin:0;overflow-wrap:anywhere}.hubActionGrid_c8d82a1a{display:grid;gap:.75rem;grid-template-columns:repeat(8,minmax(0,1fr))}.hubActionTile_c8d82a1a{align-items:center;background:linear-gradient(165deg,hsla(0,0%,100%,.08),transparent 42%),linear-gradient(145deg,#2c2c2c,#3a3a3a);border:1px solid hsla(0,0%,100%,.06);border-radius:1.05rem;box-shadow:0 10px 22px rgba(42,42,42,.14);color:#fff;cursor:pointer;display:flex;flex-direction:column;gap:.35rem;justify-content:center;min-height:5.6rem;padding:.85rem .5rem;text-align:center;transition:transform .16s ease,box-shadow .16s ease}.hubActionTile_c8d82a1a:hover{box-shadow:0 14px 28px rgba(26,26,26,.22);transform:translateY(-3px)}.hubActionTile_c8d82a1a:focus-visible{outline:3px solid #7ac143;outline-offset:2px}.hubActionIcon_c8d82a1a{align-items:center;background:rgba(122,193,67,.18);border-radius:999px;display:inline-flex;font-size:.82rem;font-weight:750;height:2.15rem;justify-content:center;width:2.15rem}.hubActionTitle_c8d82a1a{font-size:.8rem;font-weight:700;line-height:1.2}.hubActionHint_c8d82a1a{color:hsla(0,0%,100%,.62);display:none;font-size:.68rem}.hubResources_c8d82a1a{display:flex;flex-direction:column;gap:.85rem}.hubResourcesHeader_c8d82a1a{display:flex;flex-direction:column;gap:.2rem}.hubResourcesTitle_c8d82a1a{align-items:center;color:#333436;display:flex;font-size:1.05rem;font-weight:750;gap:.55rem;letter-spacing:.04em;margin:0;text-transform:uppercase}.hubResourcesTitle_c8d82a1a:before{background:#7ac143;border-radius:999px;content:\"\";height:1.15rem;width:.35rem}.hubResourcesSubtitle_c8d82a1a{color:#8a8a8a;font-size:.92rem;margin:0}.hubResourceGrid_c8d82a1a{display:grid;gap:.75rem;grid-template-columns:repeat(3,minmax(0,1fr))}.hubResourceTile_c8d82a1a{border:0;border-radius:1.05rem;box-shadow:0 14px 30px rgba(26,26,26,.16);color:#fff;cursor:pointer;display:flex;flex-direction:column;gap:.4rem;justify-content:flex-end;min-height:6rem;padding:1.1rem 1.15rem;text-align:left;transition:transform .16s ease,filter .16s ease}.hubResourceTile_c8d82a1a:hover{filter:brightness(1.05);transform:translateY(-3px)}.hubResourceTile_c8d82a1a:focus-visible{outline:3px solid #7ac143;outline-offset:2px}.hubResourceTile_c8d82a1a strong{font-size:1.05rem;font-weight:750}.hubResourceTile_c8d82a1a span{font-size:.86rem;line-height:1.4;opacity:.9}.hubTone_lime_c8d82a1a{background:linear-gradient(145deg,#6f9e2a,#8dc63f)}.hubTone_charcoal_c8d82a1a{background:linear-gradient(145deg,#2a2a2a,#3d3d3d)}.hubTone_forest_c8d82a1a{background:linear-gradient(145deg,#2f5a32,#3f7a45)}.hubTone_slate_c8d82a1a{background:linear-gradient(145deg,#3a4540,#52635a)}.hubTone_moss_c8d82a1a{background:linear-gradient(145deg,#4a6b28,#6d9538)}.hubTone_ink_c8d82a1a{background:linear-gradient(145deg,#1a1a1a,#333)}@media (min-width:1100px){.hubActionHint_c8d82a1a{display:block}}@media (max-width:1200px){.hubActionGrid_c8d82a1a{grid-template-columns:repeat(4,minmax(0,1fr))}}@media (max-width:980px){.hubResourceGrid_c8d82a1a{grid-template-columns:repeat(2,minmax(0,1fr))}}@media (max-width:800px){.hubHeroPanel_c8d82a1a{grid-template-columns:1fr}}@media (max-width:720px){.hubActionGrid_c8d82a1a{grid-template-columns:repeat(2,minmax(0,1fr))}.hubHero_c8d82a1a{padding:1.2rem 1rem 1.05rem}}@media (max-width:560px){.hubResourceGrid_c8d82a1a{grid-template-columns:1fr}}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vQzovVXNlcnMvQWJkdWwvRG93bmxvYWRzL05ldyUyMGZvbGRlciUyMCgyKS9zcGZ4L3NyYy9zaGFyZWQvdWkvcG9ydGFsLm1vZHVsZS5zY3NzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiJBQUFBLGdCQUdFLGtCQUFBLENBR0Esd0JBQUEsQ0FDQSxrQkFBQSxDQUhBLGFBQUEsQ0FIQSxZQUFBLENBSUEseUNBQUEsQ0FIQSxnQkFBQSxDQU1BLGVBQUEsQ0FHRixxQkFDRSxxQkFBQSxDQUVBLGNBQUEsQ0FDQSxXQUFBLENBRkEsVUFFQSxDQUdGLHNCQUlFLGVBQUEsQ0FDQSx5Q0FBQSxDQUpBLGVBQUEsQ0FDQSxLQUFBLENBQ0EsVUFFQSxDQUdGLHlCQUVFLGtCQUFBLENBREEsWUFBQSxDQUVBLFFBQUEsQ0FFQSxlQUFBLENBREEsaUJBQ0EsQ0FHRiwwQkFDRSxhQUFBLENBQ0EsV0FBQSxDQUdGLDBDQUNFLGdCQUFBLENBQ0Esb0JBQUEsQ0FDQSx3QkFBQSxDQUdGLDRDQUVFLGdCQUFBLENBREEsY0FDQSxDQUdGLDBCQU9FLGdDQUFBLENBSkEsa0JBQUEsQ0FGQSxZQUFBLENBQ0EsTUFBQSxDQUVBLE9BQUEsQ0FDQSxXQUFBLENBQ0EsZUFBQSxDQUVBLG9CQUFBLENBR0YsdUJBSUUsY0FBQSxDQUZBLFFBQUEsQ0FDQSxtQkFBQSxDQUVBLGFBQUEsQ0FLQSxjQUFBLENBVEEsYUFBQSxDQU1BLFlBQUEsQ0FDQSxnQkFBQSxDQUNBLGVBQUEsQ0FIQSxnQkFBQSxDQUtBLGtCQUFBLENBR0YsNkJBQ0Usa0JBQUEsQ0FDQSxhQUFBLENBR0YsNkJBQ0UsK0JBQUEsQ0FDQSxhQUFBLENBR0YsOEJBRUUsa0JBQUEsQ0FEQSxZQUFBLENBSUEsYUFBQSxDQUZBLE9BQUEsQ0FDQSxnQkFDQSxDQUdGLHdCQVdFLGtCQUFBLENBQ0Esd0JBQUEsQ0FGQSxtQkFBQSxDQUZBLGFBQUEsQ0FQQSxZQUFBLENBS0EsZ0JBQUEsQ0FDQSxlQUFBLENBTEEsZUFBQSxDQUNBLGVBQUEsQ0FNQSxnQkFBQSxDQUxBLHNCQUFBLENBQ0Esa0JBT0EsQ0FHRiwwQkFJRSxlQUFBLENBRkEsd0JBQUEsQ0FDQSxtQkFBQSxDQUVBLGFBQUEsQ0FLQSxjQUFBLENBVEEsWUFBQSxDQU1BLFlBQUEsQ0FDQSxlQUFBLENBQ0EsZUFBQSxDQUhBLGdCQUlBLENBT0YsdURBQ0UsWUFBQSxDQUdGLDJCQU1FLGFBQUEsQ0FKQSxnQkFBQSxDQUNBLGVBQUEsQ0FDQSxvQkFBQSxDQUhBLGNBQUEsQ0FJQSx3QkFDQSxDQUdGLHlCQUNFLFlBQUEsQ0FDQSxxQkFBQSxDQUNBLE9BQUEsQ0FHRiwwQkFJRSxjQUFBLENBRkEsUUFBQSxDQUNBLGtCQUFBLENBTUEsYUFBQSxDQUNBLGNBQUEsQ0FKQSxZQUFBLENBQ0EsZ0JBQUEsQ0FDQSxlQUFBLENBSEEsWUFBQSxDQUpBLGVBU0EsQ0FHRixnRUFFRSxrQkFBQSxDQUdGLGtCQUVFLGVBQUEsQ0FDQSw4QkFBQSxDQUVBLFlBQUEsQ0FDQSxxQkFBQSxDQUVBLGFBQUEsQ0FEQSxPQUFBLENBSEEsaUJBQUEsQ0FIQSxXQU9BLENBR0YsZ0JBSUUsYUFBQSxDQUhBLGlCQUFBLENBQ0EsZUFBQSxDQUNBLFFBQ0EsQ0FHRixrQkFFRSxhQUFBLENBREEsZUFBQSxDQUdBLGVBQUEsQ0FEQSxjQUNBLENBR0YsZUFLRSxrQkFBQSxDQUVBLGlCQUFBLENBSEEsYUFBQSxDQUhBLG9CQUFBLENBQ0EsZUFBQSxDQUNBLGVBQUEsQ0FLQSxrQkFBQSxDQUZBLGVBRUEsQ0FHRixpQkFHRSxjQUFBLENBREEsV0FBQSxDQUdBLGlCQUFBLENBR0EsYUFBQSxDQUNBLGNBQUEsQ0FGQSxnQkFBQSxDQURBLGVBQUEsQ0FGQSxpQkFBQSxDQUhBLGVBQUEsQ0FTQSxVQUFBLENBR0YsdUJBQ0Usa0JBQUEsQ0FHRix1QkFDRSxrQkFBQSxDQUNBLFVBQUEsQ0FHRiw2QkFDRSxrQkFBQSxDQUdGLGVBTUUsa0JBQUEsQ0FEQSxxQkFBQSxDQUpBLE1BQUEsQ0FFQSxXQUFBLENBREEsc0JBQUEsQ0FFQSxVQUVBLENBR0YscUJBSUUsZUFBQSxDQUNBLG1DQUFBLENBRkEsa0JBQUEsQ0FHQSx3Q0FBQSxDQUxBLGtCQUFBLENBQ0EsaUJBSUEsQ0FHRixrQkFNRSxhQUFBLENBSkEsZUFBQSxDQUNBLGVBQUEsQ0FDQSxvQkFBQSxDQUhBLGNBQUEsQ0FJQSx3QkFDQSxDQUdGLGdCQUtFLGFBQUEsQ0FIQSxxQ0FBQSxDQUNBLGVBQUEsQ0FDQSxxQkFBQSxDQUhBLGNBSUEsQ0FHRixtQkFFRSxhQUFBLENBQ0EsaUJBQUEsQ0FDQSxnQkFBQSxDQUhBLFFBR0EsQ0FHRixnQkFDRSxZQUFBLENBRUEsUUFBQSxDQURBLHlEQUFBLENBRUEsa0JBQUEsQ0FHRixlQUNFLCtDQUFBLENBQ0EsbUNBQUEsQ0FDQSxrQkFBQSxDQUdBLHdDQUFBLENBREEsV0FBQSxDQURBLFlBRUEsQ0FHRixvQkFLRSxhQUFBLENBSEEsbUNBQUEsQ0FDQSxlQUFBLENBQ0EscUJBQUEsQ0FIQSxRQUlBLENBR0Ysb0JBTUUsYUFBQSxDQUpBLGdCQUFBLENBQ0EsZUFBQSxDQUNBLG9CQUFBLENBSEEsY0FBQSxDQU1BLHNCQUFBLENBRkEsd0JBRUEsQ0FHRixnQkFDRSxlQUFBLENBQ0EsbUNBQUEsQ0FDQSxrQkFBQSxDQUVBLHdDQUFBLENBQ0EsV0FBQSxDQUZBLFlBRUEsQ0FHRixvQkFHRSxnQ0FBQSxDQUVBLG1DQUFBLENBREEsa0JBQUEsQ0FFQSwwQkFBQSxDQUpBLGFBQUEsQ0FEQSxVQUtBLENBR0YsZ0JBRUUsd0JBQUEsQ0FDQSxnQkFBQSxDQUNBLGVBQUEsQ0FIQSxVQUdBLENBR0Ysc0NBSUUsK0JBQUEsQ0FEQSxpQkFBQSxDQURBLGVBQUEsQ0FHQSxxQkFBQSxDQUdGLG1CQUlFLGtCQUFBLENBS0EsYUFBQSxDQUhBLGdCQUFBLENBREEsZUFBQSxDQUVBLG9CQUFBLENBTkEsZUFBQSxDQU9BLHdCQUFBLENBTkEsS0FBQSxDQVFBLGtCQUFBLENBUEEsU0FPQSxDQUdGLG1CQUNFLHNCQUFBLENBR0YsZ0JBQ0UsYUFBQSxDQUNBLGlCQUFBLENBR0YsZ0JBQ0UsYUFBQSxDQUNBLGlCQUFBLENBR0Ysa0JBQ0UsWUFBQSxDQUNBLGNBQUEsQ0FDQSxPQUFBLENBQ0Esa0JBQUEsQ0FHRix5QkFFRSxlQUFBLENBREEsd0JBQUEsQ0FHQSxtQkFBQSxDQURBLGFBQUEsQ0FLQSxjQUFBLENBREEsZ0JBQUEsQ0FEQSxlQUFBLENBREEsZ0JBR0EsQ0FHRiwrQkFFRSxrQkFBQSxDQURBLG9CQUNBLENBR0YsMkJBQ0UsNEJBQUEsQ0FDQSw4QkFBQSxDQUNBLG9CQUFBLENBR0Ysa0JBQ0UsYUFBQSxDQUNBLGlCQUFBLENBQ0EsZUFBQSxDQUdGLHFCQUlFLGVBQUEsQ0FEQSxNQUFBLENBREEsZUFBQSxDQURBLGtCQUFBLENBSUEsU0FBQSxDQUdGLHVDQUVFLGtCQUFBLENBREEsU0FDQSxDQUdGLDBDQUdFLGNBQUEsQ0FEQSxXQUFBLENBSUEsY0FBQSxDQURBLGVBQUEsQ0FEQSxlQUFBLENBSUEsZ0JBQUEsQ0FEQSxlQUNBLENBR0Ysa0JBQ0UsYUFBQSxDQUdGLHdCQUNFLGFBQUEsQ0FHRiw0REFHRSxrQkFBQSxDQURBLFVBQ0EsQ0FHRix5QkFHRSw2QkFBQSxDQUNBLFlBQUEsQ0FGQSxPQUFBLENBR0Esd0JBQUEsQ0FKQSxjQUFBLENBS0EsWUFBQSxDQUdGLGlCQUdFLGVBQUEsQ0FDQSxzQ0FBQSxDQUNBLFlBQUEsQ0FDQSxxQkFBQSxDQUpBLFdBQUEsQ0FNQSxhQUFBLENBREEsWUFBQSxDQU5BLHFCQU9BLENBR0YsdUJBRUUsc0JBQUEsQ0FEQSxZQUFBLENBR0EsUUFBQSxDQURBLDZCQUFBLENBRUEsaUJBQUEsQ0FHRixtQkFDRSxZQUFBLENBSUEsTUFBQSxDQUhBLHFCQUFBLENBQ0EsUUFBQSxDQUNBLGFBQ0EsQ0FHRixnQkFDRSxZQUFBLENBQ0EscUJBQUEsQ0FDQSxPQUFBLENBR0YscUJBR0UsYUFBQSxDQUZBLGdCQUFBLENBQ0EsZUFDQSxDQUdGLDZDQUVFLHdCQUFBLENBQ0EsaUJBQUEsQ0FLQSxxQkFBQSxDQUZBLGFBQUEsQ0FEQSxpQkFBQSxDQUVBLGVBQUEsQ0FIQSxpQkFJQSxDQUdGLHlEQUVFLHNDQUFBLENBQ0Esa0JBQUEsQ0FHRix1QkFLRSw0QkFBQSxDQUpBLFlBQUEsQ0FDQSxjQUFBLENBQ0EsT0FBQSxDQUNBLGdCQUNBLENBR0YsOEJBRUUsZUFBQSxDQURBLHdCQUFBLENBR0EsbUJBQUEsQ0FEQSxhQUFBLENBSUEsY0FBQSxDQURBLGVBQUEsQ0FEQSxpQkFFQSxDQUdGLDBDQUNFLGtCQUFBLENBQ0Esb0JBQUEsQ0FDQSxVQUFBLENBR0YsdUNBRUUsa0JBQUEsQ0FEQSxVQUNBLENBR0Ysa0JBUUUsa0JBQUEsQ0FGQSx5QkFBQSxDQUNBLGtCQUFBLENBTkEsWUFBQSxDQUNBLHFCQUFBLENBQ0EsT0FBQSxDQUNBLGFBQUEsQ0FDQSxZQUdBLENBR0Ysc0JBTUUsZUFBQSxDQUZBLHdCQUFBLENBQ0EsaUJBQUEsQ0FIQSxXQUFBLENBQ0Esa0JBQUEsQ0FGQSxVQUtBLENBR0YsbUNBQ0UsZUFBQSxDQUdGLHlCQUNFLHdCQUNFLG9CQUFBLENBR0YsMEJBQ0UsWUFBQSxDQUFBLENBSUoseUJBQ0UsMEJBQ0UsWUFBQSxDQUdGLDBCQUNFLG1CQUFBLENBR0YsMkJBSUUsNkJBQUEsQ0FIQSxhQUFBLENBRUEsT0FBQSxDQURBLGNBQUEsQ0FHQSxVQUFBLENBR0YsNEJBU0UsZUFBQSxDQUhBLFFBQUEsQ0FJQSwwQ0FBQSxDQVRBLFlBQUEsQ0FDQSxxQkFBQSxDQU1BLFlBQUEsQ0FMQSxjQUFBLENBRUEsT0FBQSxDQURBLEtBQUEsQ0FRQSwwQkFBQSxDQUNBLDhCQUFBLENBQ0EsaUJBQUEsQ0FQQSxxQkFBQSxDQUlBLFVBR0EsQ0FHRixnQ0FDRSx1QkFBQSxDQUNBLGtCQUFBLENBR0Ysb0NBQ0Usc0JBQUEsQ0FHRix3Q0FDRSxzQkFBQSxDQUdGLGVBQ0Usc0JBQUEsQ0FHRixxQkFFRSxrQkFBQSxDQURBLGlCQUNBLENBR0YseUJBQ0UsYUFBQSxDQUNBLHlCQUFBLENBQ0EsaUJBQUEsQ0FHRixpQkFDRSxVQUFBLENBQUEsQ0FJSix5QkFDRSx5QkFDRSxnQkFBQSxDQUdGLGdCQUNFLDZDQUFBLENBR0YsZ0JBRUUsZUFBQSxDQURBLGVBQ0EsQ0FHRixzQ0FFRSxnQkFBQSxDQUFBLENBTUosa0JBQ0UsWUFBQSxDQUNBLHFCQUFBLENBQ0EsV0FBQSxDQUNBLFdBQUEsQ0FDQSxVQUFBLENBR0Ysa0JBUUUsNk5BQUEsQ0FMQSxxQkFBQSxDQW1CQSx5Q0FBQSxDQWZBLHFCQUFBLENBY0EsYUFBQSxDQWhCQSxXQUFBLENBSEEsYUFBQSxDQUVBLDhCQUFBLENBSEEsaUJBQUEsQ0FLQSxVQWdCQSxDQUdGLHVCQUdFLFlBQUEsQ0FDQSxVQUFBLENBQ0EsV0FBQSxDQUpBLGlCQUFBLENBQ0EsU0FHQSxDQUdGLHlCQU1FLHdCQUFBLENBSkEsZ0JBQUEsQ0FDQSxlQUFBLENBQ0Esb0JBQUEsQ0FIQSxRQUFBLENBSUEsd0JBQ0EsQ0FHRix1QkFNRSxVQUFBLENBSkEsbUNBQUEsQ0FDQSxlQUFBLENBQ0EscUJBQUEsQ0FDQSxnQkFBQSxDQUpBLFFBS0EsQ0FHRiwwQkFLRSx5QkFBQSxDQUZBLGtDQUFBLENBQ0EsZUFBQSxDQUhBLFFBQUEsQ0FDQSxlQUdBLENBR0YsdUJBTUUsZUFBQSxDQURBLGtCQUFBLENBR0Esc0NBQUEsQ0FEQSxhQUFBLENBTkEsWUFBQSxDQUVBLGdCQUFBLENBREEsb0RBQUEsQ0FPQSxXQUFBLENBTEEsc0JBS0EsQ0FHRiw0QkFNRSxhQUFBLENBSkEsZUFBQSxDQUNBLGVBQUEsQ0FDQSxvQkFBQSxDQUhBLGlCQUFBLENBSUEsd0JBQ0EsQ0FHRiwyQkFJRSxhQUFBLENBRkEsZUFBQSxDQUNBLGdCQUFBLENBRkEsUUFBQSxDQUlBLHNCQUFBLENBR0Ysd0JBQ0UsWUFBQSxDQUVBLFVBQUEsQ0FEQSw2Q0FDQSxDQUdGLHdCQUdFLGtCQUFBLENBUUEsOEdBQUEsQ0FGQSxvQ0FBQSxDQUNBLHFCQUFBLENBTUEseUNBQUEsQ0FGQSxVQUFBLENBQ0EsY0FBQSxDQWRBLFlBQUEsQ0FDQSxxQkFBQSxDQUlBLFVBQUEsQ0FGQSxzQkFBQSxDQUdBLGlCQUFBLENBQ0Esb0JBQUEsQ0FIQSxpQkFBQSxDQVlBLG1EQUFBLENBR0YsOEJBRUUseUNBQUEsQ0FEQSwwQkFDQSxDQUdGLHNDQUNFLHlCQUFBLENBQ0Esa0JBQUEsQ0FHRix3QkFFRSxrQkFBQSxDQUtBLCtCQUFBLENBREEsbUJBQUEsQ0FMQSxtQkFBQSxDQU9BLGdCQUFBLENBQ0EsZUFBQSxDQUpBLGNBQUEsQ0FGQSxzQkFBQSxDQUNBLGFBS0EsQ0FHRix5QkFDRSxlQUFBLENBQ0EsZUFBQSxDQUNBLGVBQUEsQ0FHRix3QkFHRSx5QkFBQSxDQUZBLFlBQUEsQ0FDQSxnQkFDQSxDQUdGLHVCQUNFLFlBQUEsQ0FDQSxxQkFBQSxDQUNBLFVBQUEsQ0FHRiw2QkFDRSxZQUFBLENBQ0EscUJBQUEsQ0FDQSxTQUFBLENBR0YsNEJBR0Usa0JBQUEsQ0FNQSxhQUFBLENBUEEsWUFBQSxDQUdBLGlCQUFBLENBQ0EsZUFBQSxDQUZBLFVBQUEsQ0FHQSxvQkFBQSxDQU5BLFFBQUEsQ0FPQSx3QkFDQSxDQUdGLG1DQUtFLGtCQUFBLENBREEsbUJBQUEsQ0FIQSxVQUFBLENBRUEsY0FBQSxDQURBLFlBR0EsQ0FHRiwrQkFHRSxhQUFBLENBREEsZ0JBQUEsQ0FEQSxRQUVBLENBR0YsMEJBQ0UsWUFBQSxDQUVBLFVBQUEsQ0FEQSw2Q0FDQSxDQUdGLDBCQU9FLFFBQUEsQ0FDQSxxQkFBQSxDQUlBLHlDQUFBLENBSEEsVUFBQSxDQUVBLGNBQUEsQ0FWQSxZQUFBLENBQ0EscUJBQUEsQ0FFQSxTQUFBLENBREEsd0JBQUEsQ0FFQSxlQUFBLENBQ0Esc0JBQUEsQ0FJQSxlQUFBLENBR0EsK0NBQUEsQ0FHRixnQ0FFRSx1QkFBQSxDQURBLDBCQUNBLENBR0Ysd0NBQ0UseUJBQUEsQ0FDQSxrQkFBQSxDQUdGLGlDQUNFLGlCQUFBLENBQ0EsZUFBQSxDQUdGLCtCQUNFLGdCQUFBLENBQ0EsZUFBQSxDQUNBLFVBQUEsQ0FHRix1QkFDRSxrREFBQSxDQUdGLDJCQUNFLGtEQUFBLENBR0YseUJBQ0Usa0RBQUEsQ0FHRix3QkFDRSxrREFBQSxDQUdGLHVCQUNFLGtEQUFBLENBR0Ysc0JBQ0UsK0NBQUEsQ0FHRiwwQkFDRSx3QkFDRSxhQUFBLENBQUEsQ0FJSiwwQkFDRSx3QkFDRSw2Q0FBQSxDQUFBLENBSUoseUJBQ0UsMEJBQ0UsNkNBQUEsQ0FBQSxDQUlKLHlCQUNFLHVCQUNFLHlCQUFBLENBQUEsQ0FJSix5QkFDRSx3QkFDRSw2Q0FBQSxDQUdGLGtCQUNFLDJCQUFBLENBQUEsQ0FJSix5QkFDRSwwQkFDRSx5QkFBQSxDQUFBIiwiZmlsZSI6InBvcnRhbC5tb2R1bGUuY3NzIn0= */", true);


/***/ }),

/***/ 784:
/*!***********************************************!*\
  !*** ./lib/shared/schema/sharepointSchema.js ***!
  \***********************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SHAREPOINT_LISTS: () => (/* binding */ SHAREPOINT_LISTS),
/* harmony export */   getSharePointFieldInternalNames: () => (/* binding */ getSharePointFieldInternalNames),
/* harmony export */   getSharePointFieldLabel: () => (/* binding */ getSharePointFieldLabel),
/* harmony export */   getSharePointFields: () => (/* binding */ getSharePointFields),
/* harmony export */   getSharePointList: () => (/* binding */ getSharePointList)
/* harmony export */ });
/* unused harmony exports SHAREPOINT_SITE, getSharePointListName, getSharePointDisplayName */
/**
 * Central SharePoint schema for the PAVE Training Portal.
 *
 * Site: https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin
 *
 * Rules:
 * - Backend code must import list names and internal field names from here.
 * - Do not hard-code SharePoint list/field names in components or ad-hoc services.
 * - `fields` values are SharePoint internal names (including encoded names).
 * - `labels` are UI-friendly display names only.
 */
var SHAREPOINT_SITE = {
    url: "https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin",
    hostname: "pavetraining.sharepoint.com",
    serverRelativePath: "/sites/PaveTrainingOperationAdmin",
};
var companyFields = {
    id: "ID",
    title: "Title",
    companyNumber: "CompanyNumber",
    companyName: "CompanyName",
    companySize: "CompanySize",
    registeredAddress: "RegisteredAddress",
    companyRegNumber: "CompanyRegNumber",
    vatNo: "VATNo",
    telNo: "TelNo",
    email: "Email",
    mainContact: "MainContact",
    accountsContactName: "AccountsContactName",
    /** SharePoint internal name is `Accountsaddress` (lowercase a). */
    accountsAddress: "Accountsaddress",
    /** SharePoint internal name is `AccountsContactnumber` (lowercase n). */
    accountsContactNumber: "AccountsContactnumber",
    /** SharePoint internal name is `Accountsemail` (lowercase e). */
    accountsEmail: "Accountsemail",
    /** SharePoint internal name is `Notespricesagreed` (lowercase p/a). */
    notesPricesAgreed: "Notespricesagreed",
    companyLogo: "CompanyLogo",
    status: "Status",
};
var workforceFields = {
    id: "ID",
    candidateName: "CandidateName",
    companyName: "CompanyName",
    workforceNumber: "WorkforceNumber",
    dateOfBirth: "Dateofbirth",
    department: "Department",
    status: "Status",
    /** Live internal name is Trainingmanager (lowercase m). */
    trainingManager: "Trainingmanager",
    supervisor: "Supervisor",
    email: "Email",
    cscsNumber: "CSCSNumber",
    swqrNumber: "SWQRNumber",
    eusrNumber: "EUSRNumber",
    nporsNumbers: "NPORSNumbers",
    inHouseCertificationNumber: "InHouseCertificationNumber",
};
var trainingMatrixFields = {
    id: "ID",
    candidateName: "CandidateName",
    matrixCompany: "MatrixCompany",
    companyName: "Company_x0020_Name",
    department: "Department",
    overallStatus: "OverallStatus",
    needsReview: "NeedsReview",
    matrixNotes: "MatrixNotes",
    nextExpiryDate: "NextExpiryDate",
    n001Expiry: "N001Expiry",
    n003Expiry: "N003Expiry",
    n004Expiry: "N004Expiry",
    n010Expiry: "N010Expiry",
    n020Expiry: "N020Expiry",
    n021Expiry: "N021Expiry",
    n027Expiry: "N027Expiry",
    n100Expiry: "N100Expiry",
};
var nporsRegisterFields = {
    id: "ID",
    candidateName: "CandidateName",
    companyName: "CompanyName",
    nporsNumber: "NPORSNumber",
    trainingDate: "TrainingDate",
    trainingAddress: "TrainingAddress",
    noviceOrEwt: "NoviceorEwt",
    nporsCategory: "NPORSCategory",
    expiry: "Expiry",
    trainingOutcome: "TrainingOutcome",
    outcomeDate: "OutcomeDate",
    assessorTrainer: "AssessorTrainer",
    customerVisible: "CustomerVisible",
    outcomeNotes: "OutcomeNotes",
    notes: "Notes",
};
var eusrRegisterFields = {
    id: "ID",
    candidateName: "CandidateName",
    companyName: "CompanyName",
    eusrNumber: "EUSRNumber",
    eusrCategory: "EusrCategory",
    cardType: "CardType",
    trainingDate: "TrainingDate",
    trainingAddress: "TrainingAddress",
    expiry: "Expiry",
    trainingOutcome: "TrainingOutcome",
    outcomeDate: "OutcomeDate",
    assessorTrainer: "AssessorTrainer",
    customerVisible: "CustomerVisible",
    outcomeNotes: "OutcomeNotes",
    notes: "Notes",
};
var nrswaRegisterFields = {
    id: "ID",
    candidateName: "CandidateName",
    companyName: "CompanyName",
    swqrNumber: "SWQRNumber",
    streetworksCategory: "StreetworksCategory",
    course: "Course",
    trainingDate: "TrainingDate",
    trainingAddress: "TrainingAddress",
    expiryDate: "Expirydate",
    trainingOutcome: "TrainingOutcome",
    outcomeDate: "OutcomeDate",
    assessorTrainer: "AssessorTrainer",
    customerVisible: "CustomerVisible",
    outcomeNotes: "OutcomeNotes",
};
var inHouseCertificatesFields = {
    id: "ID",
    candidateName: "CandidateName",
    companyName: "CompanyName",
    certificateCategory: "CertificateCategory",
    courseCategory: "CourseCategory",
    courseDate: "CourseDate",
    trainingAddress: "TrainingAddress",
    expiryDate: "ExpiryDate",
    trainingOutcome: "TrainingOutcome",
    outcomeDate: "OutcomeDate",
    assessorTrainer: "AssessorTrainer",
    customerVisible: "CustomerVisible",
    outcomeNotes: "OutcomeNotes",
    notes: "Notes",
};
var nvqRegisterFields = {
    id: "ID",
    candidateName: "CandidateName",
    nvqCompany: "NVQCompany",
    companyName: "Company_x0020_Name",
    nvqTitle: "NvqTitle",
    boltonNvq: "BoltonNvq",
    dateRegistered: "DateRegistered",
    dateInductionBooked: "DateinductionBooked",
    stageOfNvq: "StageofNvq",
    customerUpdateNotes: "CustomerUpdateNotes",
    completedDate: "CompletedDate",
    customerVisible: "CustomerVisible",
    trainingOutcome: "TrainingOutcome",
    outcomeDate: "OutcomeDate",
    assessorTrainer: "AssessorTrainer",
    outcomeNotes: "OutcomeNotes",
};
var customerDocumentsFields = {
    id: "ID",
    title: "Title",
    company: "Company",
    candidate: "Candidate",
    documentType: "DocumentType",
    customerVisible: "CustomerVisible",
    notificationSent: "NotificationSent",
    fileRef: "FileRef",
    fileLeafRef: "FileLeafRef",
    fsObjType: "FSObjType",
    modified: "Modified",
    editor: "Editor",
};
var eventsFields = {
    id: "ID",
    title: "Title",
    /** Lookup to Company List — do not use a legacy `Company` field. */
    eventCompany: "EventCompany",
    /** Graph companion for EventCompany lookup writes. */
    eventCompanyLookupId: "EventCompanyLookupId",
    customerVisible: "Customer_x0020_Visible",
    trainingAddress: "TrainingAddress",
    eventDate: "EventDate",
    endDate: "EndDate",
    description: "Description",
    location: "Location",
    outlookEventId: "OutlookEventId",
    outlookCalendarId: "OutlookCalendarId",
    outlookICalUid: "OutlookICalUid",
    syncStatus: "SyncStatus",
    syncDirection: "SyncDirection",
    lastSyncedAt: "LastSyncedAt",
    lastSyncSource: "LastSyncSource",
    syncError: "SyncError",
    doNotSync: "DoNotSync",
};
var offersPromotionsFields = {
    id: "ID",
    title: "Title",
    category: "Category",
    customerVisible: "CustomerVisible",
    startDate: "StartDate",
    endDate: "EndDate",
    /** Live SharePoint field is ShortDescription (not Description). */
    shortDescription: "ShortDescription",
    status: "Status",
};
var permissionsFields = {
    id: "ID",
    userEmail: "UserEmail",
    roleType: "RoleType",
    status: "Status",
    company: "Company",
    /** Graph expand companion for lookup field `Company`. */
    companyLookupId: "CompanyLookupId",
    accessScope: "AccessScope",
    /** Person/display name — used for CandidateOnly and Supervisor matching. */
    name: "Name",
    /** Multi-choice department scopes for Supervisor. */
    departments: "Departments",
    /** Lookup multi to Departments list. */
    departmentsAllowed: "DepartmentsAllowed",
    canView: "CanView",
    canDownload: "CanDownload",
    canEdit: "CanEdit",
};
var trainingCourseCategoriesFields = {
    id: "ID",
    title: "Title",
    categoryCode: "CategoryCode",
    courseName: "CourseName",
    courseType: "CourseType",
    source: "Source",
    active: "Active",
    customerVisible: "CustomerVisible",
    displayOrder: "DisplayOrder",
    notes: "Notes",
};
var trainingManagerLogsFields = {
    id: "ID",
    title: "Title",
    /** Encoded internal name on live list. */
    userEmail: "User_x0020_Email",
    listName: "ListName",
    itemsId: "ItemsId",
    areaViewed: "Area_x0020_Viewed",
    timestamp: "Timestamp",
    notes: "Notes",
    company: "Company",
    role: "Role",
};
var SHAREPOINT_LISTS = {
    company: {
        key: "company",
        listName: "Company List",
        displayName: "Company",
        listIdEnvVar: "SHAREPOINT_COMPANY_LIST_ID",
        fields: companyFields,
        labels: {
            id: "ID",
            title: "Title",
            companyNumber: "Company number",
            companyName: "Company name",
            companySize: "Company size",
            registeredAddress: "Registered address",
            companyRegNumber: "Company reg number",
            vatNo: "VAT no",
            telNo: "Tel no",
            email: "Email",
            mainContact: "Main contact",
            accountsContactName: "Accounts contact name",
            accountsAddress: "Accounts address",
            accountsContactNumber: "Accounts contact number",
            accountsEmail: "Accounts email",
            notesPricesAgreed: "Notes / prices agreed",
            companyLogo: "Company logo",
            status: "Status",
        },
    },
    workforce: {
        key: "workforce",
        listName: "Workforce List",
        displayName: "Workforce",
        listIdEnvVar: "SHAREPOINT_WORKFORCE_LIST_ID",
        fields: workforceFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            companyName: "Company name",
            workforceNumber: "Workforce number",
            dateOfBirth: "Date of birth",
            department: "Department",
            status: "Status",
            trainingManager: "Training manager",
            supervisor: "Supervisor",
            email: "Email",
            cscsNumber: "CSCS number",
            swqrNumber: "SWQR number",
            eusrNumber: "EUSR number",
            nporsNumbers: "NPORS numbers",
            inHouseCertificationNumber: "In-house certification number",
        },
    },
    trainingMatrix: {
        key: "trainingMatrix",
        listName: "Training Matrix",
        displayName: "Training Matrix",
        listIdEnvVar: "SHAREPOINT_TRAINING_MATRIX_LIST_ID",
        fields: trainingMatrixFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            matrixCompany: "Matrix company",
            companyName: "Company name",
            department: "Department",
            overallStatus: "Overall status",
            needsReview: "Needs review",
            matrixNotes: "Matrix notes",
            nextExpiryDate: "Next expiry date",
            n001Expiry: "N001 expiry",
            n003Expiry: "N003 expiry",
            n004Expiry: "N004 expiry",
            n010Expiry: "N010 expiry",
            n020Expiry: "N020 expiry",
            n021Expiry: "N021 expiry",
            n027Expiry: "N027 expiry",
            n100Expiry: "N100 expiry",
        },
    },
    nporsRegister: {
        key: "nporsRegister",
        listName: "NPORS Register",
        displayName: "NPORS",
        listIdEnvVar: "SHAREPOINT_NPORS_REGISTER_LIST_ID",
        fields: nporsRegisterFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            companyName: "Company name",
            nporsNumber: "NPORS number",
            trainingDate: "Training date",
            trainingAddress: "Training address",
            noviceOrEwt: "Novice or EWT",
            nporsCategory: "NPORS category",
            expiry: "Expiry",
            trainingOutcome: "Training outcome",
            outcomeDate: "Outcome date",
            assessorTrainer: "Assessor / trainer",
            customerVisible: "Customer visible",
            outcomeNotes: "Outcome notes",
            notes: "Notes",
        },
    },
    eusrRegister: {
        key: "eusrRegister",
        listName: "EUSR Register",
        displayName: "EUSR",
        listIdEnvVar: "SHAREPOINT_EUSR_REGISTER_LIST_ID",
        fields: eusrRegisterFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            companyName: "Company name",
            eusrNumber: "EUSR number",
            eusrCategory: "EUSR category",
            cardType: "Card type",
            trainingDate: "Training date",
            trainingAddress: "Training address",
            expiry: "Expiry",
            trainingOutcome: "Training outcome",
            outcomeDate: "Outcome date",
            assessorTrainer: "Assessor / trainer",
            customerVisible: "Customer visible",
            outcomeNotes: "Outcome notes",
            notes: "Notes",
        },
    },
    nrswaRegister: {
        key: "nrswaRegister",
        listName: "NRSWA Register",
        displayName: "Streetworks Training",
        listIdEnvVar: "SHAREPOINT_NRSWA_REGISTER_LIST_ID",
        fields: nrswaRegisterFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            companyName: "Company name",
            swqrNumber: "SWQR number",
            streetworksCategory: "Streetworks category",
            course: "Course",
            trainingDate: "Training date",
            trainingAddress: "Training address",
            expiryDate: "Expiry date",
            trainingOutcome: "Training outcome",
            outcomeDate: "Outcome date",
            assessorTrainer: "Assessor / trainer",
            customerVisible: "Customer visible",
            outcomeNotes: "Outcome notes",
        },
    },
    inHouseCertificates: {
        key: "inHouseCertificates",
        listName: "In-House Certificates Register",
        displayName: "In-House Certificates",
        listIdEnvVar: "SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID",
        fields: inHouseCertificatesFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            companyName: "Company name",
            certificateCategory: "Certificate category",
            courseCategory: "Course category",
            courseDate: "Course date",
            trainingAddress: "Training address",
            expiryDate: "Expiry date",
            trainingOutcome: "Training outcome",
            outcomeDate: "Outcome date",
            assessorTrainer: "Assessor / trainer",
            customerVisible: "Customer visible",
            outcomeNotes: "Outcome notes",
            notes: "Notes",
        },
    },
    nvqRegister: {
        key: "nvqRegister",
        listName: "NVQ Register",
        displayName: "NVQ",
        listIdEnvVar: "SHAREPOINT_NVQ_REGISTER_LIST_ID",
        fields: nvqRegisterFields,
        labels: {
            id: "ID",
            candidateName: "Candidate name",
            nvqCompany: "NVQ company",
            companyName: "Company name",
            nvqTitle: "NVQ title",
            boltonNvq: "Bolton NVQ",
            dateRegistered: "Date registered",
            dateInductionBooked: "Date induction booked",
            stageOfNvq: "Stage of NVQ",
            customerUpdateNotes: "Customer update notes",
            completedDate: "Completed date",
            customerVisible: "Customer visible",
            trainingOutcome: "Training outcome",
            outcomeDate: "Outcome date",
            assessorTrainer: "Assessor / trainer",
            outcomeNotes: "Outcome notes",
        },
    },
    customerDocuments: {
        key: "customerDocuments",
        listName: "Customer Documents",
        displayName: "Customer Documents",
        listIdEnvVar: "SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID",
        fields: customerDocumentsFields,
        labels: {
            id: "ID",
            title: "Title",
            company: "Company",
            candidate: "Candidate",
            documentType: "Document type",
            customerVisible: "Customer visible",
            notificationSent: "Notification sent",
            fileRef: "File path",
            fileLeafRef: "File name",
            fsObjType: "Object type",
            modified: "Modified date",
            editor: "Modified by",
        },
    },
    events: {
        key: "events",
        listName: "Events",
        displayName: "Events",
        listIdEnvVar: "SHAREPOINT_EVENTS_LIST_ID",
        fields: eventsFields,
        labels: {
            id: "ID",
            title: "Title",
            eventCompany: "Company",
            eventCompanyLookupId: "Company lookup ID",
            customerVisible: "Customer visible",
            trainingAddress: "Training address",
            eventDate: "Start date / time",
            endDate: "End date / time",
            description: "Description",
            location: "Location",
            outlookEventId: "Outlook event ID",
            outlookCalendarId: "Outlook calendar ID",
            outlookICalUid: "Outlook iCal UID",
            syncStatus: "Sync status",
            syncDirection: "Sync direction",
            lastSyncedAt: "Last synced at",
            lastSyncSource: "Last sync source",
            syncError: "Sync error",
            doNotSync: "Do not sync",
        },
    },
    offersPromotions: {
        key: "offersPromotions",
        listName: "Offers / Promotions",
        displayName: "Offers / Promotions",
        /** GUID avoids REST 404 — slash in title breaks getbytitle URLs. */
        listId: "8e887fc7-0404-47e3-977c-e6e24e0b85c6",
        listIdEnvVar: "SHAREPOINT_OFFERS_PROMOTIONS_LIST_ID",
        fields: offersPromotionsFields,
        labels: {
            id: "ID",
            title: "Title",
            category: "Category",
            customerVisible: "Customer visible",
            startDate: "Start date",
            endDate: "End date",
            shortDescription: "Short description",
            status: "Status",
        },
    },
    permissions: {
        key: "permissions",
        listName: "Permissions List",
        displayName: "Permissions",
        listIdEnvVar: "SHAREPOINT_PERMISSIONS_LIST_ID",
        fields: permissionsFields,
        labels: {
            id: "ID",
            userEmail: "User email",
            roleType: "Role type",
            status: "Status",
            company: "Company",
            companyLookupId: "Company lookup ID",
            accessScope: "Access scope",
            name: "Name",
            departments: "Departments",
            departmentsAllowed: "Departments allowed",
            canView: "Can view",
            canDownload: "Can download",
            canEdit: "Can edit",
        },
    },
    trainingCourseCategories: {
        key: "trainingCourseCategories",
        listName: "Training Course Categories",
        displayName: "Training Course Categories",
        listIdEnvVar: "SHAREPOINT_TRAINING_COURSE_CATEGORIES_LIST_ID",
        fields: trainingCourseCategoriesFields,
        labels: {
            id: "ID",
            title: "Title",
            categoryCode: "Category code",
            courseName: "Course name",
            courseType: "Course type",
            source: "Source",
            active: "Active",
            customerVisible: "Customer visible",
            displayOrder: "Display order",
            notes: "Notes",
        },
    },
    trainingManagerLogs: {
        key: "trainingManagerLogs",
        listName: "Training Manager Logs",
        displayName: "Training Manager Logs",
        listIdEnvVar: "SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID",
        fields: trainingManagerLogsFields,
        labels: {
            id: "ID",
            title: "Title",
            userEmail: "User email",
            listName: "List name",
            itemsId: "Items ID",
            areaViewed: "Area viewed",
            timestamp: "Timestamp",
            notes: "Notes",
            company: "Company",
            role: "Role",
        },
    },
};
function getSharePointList(listKey) {
    return SHAREPOINT_LISTS[listKey];
}
function getSharePointListName(listKey) {
    return SHAREPOINT_LISTS[listKey].listName;
}
function getSharePointDisplayName(listKey) {
    return SHAREPOINT_LISTS[listKey].displayName;
}
function getSharePointFields(listKey) {
    return SHAREPOINT_LISTS[listKey].fields;
}
function getSharePointFieldInternalNames(listKey) {
    var fields = SHAREPOINT_LISTS[listKey].fields;
    var names = [];
    for (var key in fields) {
        if (Object.prototype.hasOwnProperty.call(fields, key)) {
            names.push(fields[key]);
        }
    }
    return names;
}
function getSharePointFieldLabel(listKey, fieldKey) {
    var _a;
    var labels = SHAREPOINT_LISTS[listKey].labels;
    return (_a = labels[fieldKey]) !== null && _a !== void 0 ? _a : fieldKey;
}


/***/ }),

/***/ 827:
/*!***************************************************!*\
  !*** ./lib/shared/services/companyLogoService.js ***!
  \***************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   thumbnailPreviewUrl: () => (/* binding */ thumbnailPreviewUrl),
/* harmony export */   uploadAndSetListImage: () => (/* binding */ uploadAndSetListImage)
/* harmony export */ });
/* unused harmony exports parseThumbnailField, isThumbnailJsonString */
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-http */ 272);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
/* harmony import */ var _sharePointListService__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./sharePointListService */ 161);
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};



function escapeOData(value) {
    return value.replace(/'/g, "''");
}
function siteOrigin(webUrl) {
    try {
        return new URL(webUrl).origin;
    }
    catch (_a) {
        var i = webUrl.indexOf("/", webUrl.indexOf("//") + 2);
        return i > 0 ? webUrl.substring(0, i) : webUrl;
    }
}
/**
 * Parse SharePoint Image/Thumbnail column value (object or JSON string).
 */
function parseThumbnailField(value) {
    if (value === null || value === undefined || value === "")
        return null;
    var parsed = value;
    if (typeof value === "string") {
        var trimmed = value.trim();
        if (!trimmed)
            return null;
        if (trimmed.charAt(0) === "{") {
            try {
                parsed = JSON.parse(trimmed);
            }
            catch (_a) {
                return { serverRelativeUrl: trimmed };
            }
        }
        else if (trimmed.charAt(0) === "/") {
            return { serverRelativeUrl: trimmed };
        }
        else {
            return null;
        }
    }
    if (typeof parsed === "object" && parsed !== null) {
        var rec = parsed;
        return {
            serverRelativeUrl: typeof rec.serverRelativeUrl === "string"
                ? rec.serverRelativeUrl
                : undefined,
            fileName: typeof rec.fileName === "string" ? rec.fileName : undefined,
            id: typeof rec.id === "string" ? rec.id : undefined,
        };
    }
    return null;
}
function thumbnailPreviewUrl(webUrl, value) {
    var meta = parseThumbnailField(value);
    if (!meta || !meta.serverRelativeUrl)
        return null;
    return siteOrigin(webUrl) + meta.serverRelativeUrl;
}
function getListId(client, listKey) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, url, response, text, json;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    url = client.webUrl +
                        "/_api/web/lists/getbytitle('" +
                        escapeOData(listName) +
                        "')?$select=Id";
                    return [4 /*yield*/, client.spHttpClient.get(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, {
                            headers: {
                                Accept: "application/json;odata=nometadata",
                                "odata-version": "",
                            },
                        })];
                case 1:
                    response = _a.sent();
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    text = _a.sent();
                    throw new Error('Failed to resolve list id for "' + listName + '": ' + text);
                case 3: return [4 /*yield*/, response.json()];
                case 4:
                    json = (_a.sent());
                    if (!json.Id) {
                        throw new Error('List id missing for "' + listName + '".');
                    }
                    return [2 /*return*/, String(json.Id)];
            }
        });
    });
}
function safeImageFileName(file) {
    var raw = (file.name || "logo.png").replace(/[^\w.\-]+/g, "_");
    var parts = raw.split(".");
    var ext = parts.length > 1 ? parts[parts.length - 1].toLowerCase() : "png";
    var base = parts.length > 1 ? parts.slice(0, -1).join(".") : parts[0] || "logo";
    return Date.now() + "-" + base.substring(0, 40) + "." + ext;
}
/**
 * Upload an image into the list's Site Assets folder via SharePoint UploadImage,
 * then set the Thumbnail/Image column (CompanyLogo) on the item.
 */
function uploadAndSetListImage(client, listKey, itemId, fieldInternalName, file) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, listId, imageName, uploadUrl, buffer, response, text, uploaded, meta;
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!file || !file.size) {
                        throw new Error("Please choose an image file to upload.");
                    }
                    if (file.size > 10 * 1024 * 1024) {
                        throw new Error("Image must be 10 MB or smaller.");
                    }
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    return [4 /*yield*/, getListId(client, listKey)];
                case 1:
                    listId = _b.sent();
                    imageName = safeImageFileName(file);
                    uploadUrl = client.webUrl +
                        "/_api/web/uploadimage(listTitle=@a1,imageName=@a2,listId=@a3,itemId=@a4)" +
                        "?@a1='" +
                        escapeOData(listName) +
                        "'&@a2='" +
                        escapeOData(imageName) +
                        "'&@a3='" +
                        listId +
                        "'&@a4=" +
                        itemId;
                    return [4 /*yield*/, file.arrayBuffer()];
                case 2:
                    buffer = _b.sent();
                    return [4 /*yield*/, client.spHttpClient.post(uploadUrl, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, {
                            headers: {
                                Accept: "application/json;odata=nometadata",
                                "Content-Type": "application/octet-stream",
                                "odata-version": "",
                            },
                            body: buffer,
                        })];
                case 3:
                    response = _b.sent();
                    if (!!response.ok) return [3 /*break*/, 5];
                    return [4 /*yield*/, response.text()];
                case 4:
                    text = _b.sent();
                    throw new Error("Image upload failed (" + response.status + "): " + text);
                case 5: return [4 /*yield*/, response.json()];
                case 6:
                    uploaded = (_b.sent());
                    if (!uploaded.ServerRelativeUrl) {
                        throw new Error("Image upload succeeded but no ServerRelativeUrl returned.");
                    }
                    meta = {
                        type: "thumbnail",
                        fileName: uploaded.Name || imageName,
                        fieldName: fieldInternalName,
                        serverUrl: siteOrigin(client.webUrl),
                        serverRelativeUrl: uploaded.ServerRelativeUrl,
                        id: uploaded.UniqueId,
                    };
                    // Image columns must be a JSON *string*, not a nested object.
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.updateListItem)(client, listKey, itemId, (_a = {},
                            _a[fieldInternalName] = JSON.stringify(meta),
                            _a))];
                case 7:
                    // Image columns must be a JSON *string*, not a nested object.
                    _b.sent();
                    return [2 /*return*/, meta];
            }
        });
    });
}
/** Allow CompanyLogo when it is already a stringified thumbnail JSON. */
function isThumbnailJsonString(value) {
    return typeof value === "string" && value.trim().charAt(0) === "{";
}


/***/ }),

/***/ 410:
/*!******************************************************!*\
  !*** ./lib/shared/services/customerAccessService.js ***!
  \******************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   documentRowsToPortalTable: () => (/* binding */ documentRowsToPortalTable),
/* harmony export */   filterPortalRowsByAccess: () => (/* binding */ filterPortalRowsByAccess),
/* harmony export */   loadCustomerDocuments: () => (/* binding */ loadCustomerDocuments)
/* harmony export */ });
/* unused harmony export candidateRecordAllowed */
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
/* harmony import */ var _permissionService__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./permissionService */ 942);
/* harmony import */ var _sharePointListService__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./sharePointListService */ 161);
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};



var docs = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("customerDocuments");
var workforce = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("workforce");
function cell(value) {
    var _a;
    return (_a = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value)) !== null && _a !== void 0 ? _a : "";
}
function lookupId(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
        return String(value);
    }
    if (typeof value === "string" && /^\d+$/.test(value.trim())) {
        return value.trim();
    }
    if (value && typeof value === "object") {
        var id = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value.LookupId) ||
            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value.Id) ||
            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value.id);
        if (id && /^\d+$/.test(id))
            return id;
    }
    return null;
}
function lookupDisplay(value) {
    if (typeof value === "string") {
        var trimmed = value.trim();
        if (!trimmed || /^\d+$/.test(trimmed))
            return null;
        return trimmed;
    }
    if (value && typeof value === "object") {
        return ((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value.LookupValue) ||
            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value.Title) ||
            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asString)(value.CandidateName) ||
            null);
    }
    return null;
}
function isSharePointFile(fields) {
    var fs = fields[docs.fsObjType];
    if (fs === 1 || fs === "1")
        return false;
    if (fs === 0 || fs === "0")
        return true;
    var leaf = cell(fields[docs.fileLeafRef]);
    if (!leaf)
        return false;
    return leaf.indexOf(".") >= 0;
}
function formatDate(iso) {
    if (!iso)
        return "";
    var d = new Date(iso);
    if (isNaN(d.getTime()))
        return iso;
    return d.toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
    });
}
function departmentsFromWorkforce(fields) {
    var raw = fields[workforce.department];
    if (Array.isArray(raw)) {
        return raw.map(function (x) { return cell(x); }).filter(Boolean);
    }
    var single = cell(raw);
    if (!single)
        return [];
    return single.split(/;|,/).map(function (p) { return p.trim(); }).filter(Boolean);
}
function supervisorMatches(fields, permission) {
    var supervisor = fields[workforce.supervisor];
    var display = lookupDisplay(supervisor);
    var id = lookupId(supervisor) || cell(fields.SupervisorId);
    if (id && id === permission.permissionItemId)
        return true;
    if (display &&
        permission.candidateScopeName &&
        display.toLowerCase() === permission.candidateScopeName.toLowerCase()) {
        return true;
    }
    if (display && display.toLowerCase() === permission.userEmail.toLowerCase()) {
        return true;
    }
    if (display &&
        permission.roleLabel &&
        display.toLowerCase() === permission.roleLabel.toLowerCase()) {
        return true;
    }
    return false;
}
/**
 * Whether a workforce / matrix-style row is visible under the user's access scope.
 */
function candidateRecordAllowed(fields, permission, options) {
    var _a;
    var scope = permission.normalizedAccessScope;
    if (scope === "Company" || scope === "All")
        return true;
    var nameField = (options === null || options === void 0 ? void 0 : options.candidateNameField) || workforce.candidateName;
    var deptField = (options === null || options === void 0 ? void 0 : options.departmentField) || workforce.department;
    var candidateName = cell(fields[nameField]);
    if (scope === "CandidateOnly") {
        var email = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.normalizeSharePointUserEmail)(cell(fields[workforce.email]));
        if (email && email === permission.userEmail)
            return true;
        if (permission.candidateScopeName &&
            candidateName &&
            candidateName.toLowerCase() ===
                permission.candidateScopeName.toLowerCase()) {
            return true;
        }
        return false;
    }
    // Department / AssignedCandidates
    if (supervisorMatches(fields, permission))
        return true;
    if (permission.departmentScopes.length > 0) {
        var depts = departmentsFromWorkforce(__assign(__assign({}, fields), (_a = {}, _a[workforce.department] = fields[deptField], _a)));
        for (var i = 0; i < depts.length; i++) {
            for (var j = 0; j < permission.departmentScopes.length; j++) {
                if (depts[i].toLowerCase() ===
                    permission.departmentScopes[j].toLowerCase()) {
                    return true;
                }
            }
        }
        return false;
    }
    // No department scopes configured — fall back to supervisor assignment only.
    return supervisorMatches(fields, permission);
}
function filterPortalRowsByAccess(rows, permission, options) {
    if (permission.normalizedAccessScope === "Company" ||
        permission.normalizedAccessScope === "All") {
        return rows;
    }
    return rows.filter(function (row) {
        return candidateRecordAllowed(row.fields || {}, permission, options);
    });
}
function resolveCandidateName(client, candidateValue, candidateIdHint, cache) {
    return __awaiter(this, void 0, void 0, function () {
        var display, id, item, name_1, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    display = lookupDisplay(candidateValue);
                    if (display)
                        return [2 /*return*/, display];
                    id = candidateIdHint || lookupId(candidateValue);
                    if (!id)
                        return [2 /*return*/, "—"];
                    if (cache[id])
                        return [2 /*return*/, cache[id]];
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.getListItem)(client, "workforce", id)];
                case 2:
                    item = _b.sent();
                    name_1 = (item && cell(item.fields[workforce.candidateName])) ||
                        (item && cell(item.fields.Title)) ||
                        "";
                    cache[id] = name_1 || "—";
                    return [2 /*return*/, cache[id]];
                case 3:
                    _a = _b.sent();
                    cache[id] = "—";
                    return [2 /*return*/, "—"];
                case 4: return [2 /*return*/];
            }
        });
    });
}
function buildFileUrl(webUrl, fields) {
    var abs = cell(fields.EncodedAbsUrl);
    if (abs)
        return abs;
    var fileRef = cell(fields[docs.fileRef]);
    if (!fileRef)
        return null;
    if (/^https?:\/\//i.test(fileRef))
        return fileRef;
    var root = webUrl.replace(/\/$/, "");
    if (fileRef.indexOf("/") === 0)
        return root.replace(/^(https?:\/\/[^/]+).*$/i, "$1") + fileRef;
    return root + "/" + fileRef;
}
/**
 * Customer Documents — company + CustomerVisible + files only,
 * with candidate names resolved and role scope applied.
 */
function loadCustomerDocuments(client, permission) {
    return __awaiter(this, void 0, void 0, function () {
        var companyId, companyName, items, nameCache, allowedCandidateIds, workforceIndex, needsScopeFilter, i, wf, n, rows, i, item, fields, companyDisplay, candidateRaw, candidateId, candidateName, name_2, modified, viewUrl, canDownload;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    companyId = permission.companyId;
                    companyName = permission.companyDisplayName || "";
                    if (!companyId || companyId === "0")
                        return [2 /*return*/, []];
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.getListItems)(client, "customerDocuments", {
                            filter: ["CompanyId eq " + Number(companyId), docs.customerVisible + " eq 1"].join(" and "),
                            top: 5000,
                            maxItems: 5000,
                        })];
                case 1:
                    items = _a.sent();
                    nameCache = {};
                    allowedCandidateIds = new Set();
                    workforceIndex = null;
                    needsScopeFilter = permission.normalizedAccessScope !== "Company" &&
                        permission.normalizedAccessScope !== "All";
                    if (!needsScopeFilter) return [3 /*break*/, 3];
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.getListItems)(client, "workforce", {
                            filter: "CompanyNameId eq " + Number(companyId),
                            top: 5000,
                            maxItems: 5000,
                        }).catch(function () { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                return [2 /*return*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.getListItems)(client, "workforce", {
                                        top: 5000,
                                        maxItems: 5000,
                                    })];
                            });
                        }); })];
                case 2:
                    workforceIndex = _a.sent();
                    for (i = 0; i < workforceIndex.length; i++) {
                        wf = workforceIndex[i];
                        if (candidateRecordAllowed(wf.fields, permission)) {
                            allowedCandidateIds.add(wf.id);
                            n = cell(wf.fields[workforce.candidateName]);
                            if (n)
                                nameCache[wf.id] = n;
                        }
                    }
                    _a.label = 3;
                case 3:
                    rows = [];
                    i = 0;
                    _a.label = 4;
                case 4:
                    if (!(i < items.length)) return [3 /*break*/, 7];
                    item = items[i];
                    fields = item.fields;
                    if (!isSharePointFile(fields))
                        return [3 /*break*/, 6];
                    if (!(0,_sharePointListService__WEBPACK_IMPORTED_MODULE_2__.asBoolean)(fields[docs.customerVisible]))
                        return [3 /*break*/, 6];
                    companyDisplay = lookupDisplay(fields[docs.company]) ||
                        cell(fields.Company) ||
                        companyName;
                    if (!companyDisplay && !companyId)
                        return [3 /*break*/, 6];
                    candidateRaw = fields[docs.candidate];
                    candidateId = lookupId(candidateRaw) || cell(fields.CandidateId) || null;
                    if (needsScopeFilter) {
                        // Company-level docs (no candidate) — Training Manager only; supervisors skip.
                        if (!candidateId) {
                            if (permission.customerRole !== "TrainingManager")
                                return [3 /*break*/, 6];
                        }
                        else if (!allowedCandidateIds.has(candidateId)) {
                            return [3 /*break*/, 6];
                        }
                    }
                    return [4 /*yield*/, resolveCandidateName(client, candidateRaw, candidateId, nameCache)];
                case 5:
                    candidateName = _a.sent();
                    name_2 = cell(fields[docs.fileLeafRef]) ||
                        cell(fields[docs.title]) ||
                        "Document";
                    modified = cell(fields[docs.modified]) ||
                        cell(fields.Modified) ||
                        "";
                    viewUrl = buildFileUrl(client.webUrl, fields);
                    canDownload = permission.canDownload === true;
                    rows.push({
                        id: item.id,
                        name: name_2,
                        documentType: cell(fields[docs.documentType]) || "Document",
                        candidateName: candidateName,
                        companyName: companyDisplay || companyName,
                        modifiedDate: formatDate(modified),
                        viewUrl: viewUrl,
                        downloadUrl: canDownload && viewUrl ? viewUrl : null,
                        canDownload: canDownload,
                    });
                    _a.label = 6;
                case 6:
                    i++;
                    return [3 /*break*/, 4];
                case 7: return [2 /*return*/, rows];
            }
        });
    });
}
function documentRowsToPortalTable(docsRows) {
    return docsRows.map(function (row) { return ({
        id: row.id,
        cells: [
            row.name,
            row.documentType,
            row.candidateName,
            row.modifiedDate,
            row.viewUrl || "",
            row.canDownload && row.downloadUrl ? row.downloadUrl : "",
        ],
        fields: {
            __docViewUrl: row.viewUrl,
            __docDownloadUrl: row.downloadUrl,
            __docCanDownload: row.canDownload,
        },
    }); });
}



/***/ }),

/***/ 237:
/*!**********************************************!*\
  !*** ./lib/shared/services/exportService.js ***!
  \**********************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   exportTableAsCsv: () => (/* binding */ exportTableAsCsv),
/* harmony export */   exportTableAsExcel: () => (/* binding */ exportTableAsExcel)
/* harmony export */ });
/**
 * Client-side CSV / Excel downloads (no extra packages).
 */
function escapeCsvCell(value) {
    var needsQuotes = value.indexOf(",") >= 0 ||
        value.indexOf('"') >= 0 ||
        value.indexOf("\n") >= 0 ||
        value.indexOf("\r") >= 0;
    var escaped = value.replace(/"/g, '""');
    return needsQuotes ? '"' + escaped + '"' : escaped;
}
function triggerDownload(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
function stamp() {
    var d = new Date();
    var p = function (n) { return (n < 10 ? "0" + n : String(n)); };
    return (d.getFullYear() +
        p(d.getMonth() + 1) +
        p(d.getDate()) +
        "-" +
        p(d.getHours()) +
        p(d.getMinutes()));
}
function exportTableAsCsv(title, headers, rows) {
    var lines = [];
    lines.push(headers.map(escapeCsvCell).join(","));
    for (var i = 0; i < rows.length; i++) {
        lines.push(rows[i].map(function (c) { return escapeCsvCell(c == null ? "" : String(c)); }).join(","));
    }
    // UTF-8 BOM so Excel opens accents correctly
    var csv = "\uFEFF" + lines.join("\r\n");
    var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    var safe = title.replace(/[^\w\-]+/g, "_").replace(/^_+|_+$/g, "") || "export";
    triggerDownload(blob, safe + "-" + stamp() + ".csv");
}
/**
 * Excel-compatible .xls via HTML table (opens directly in Microsoft Excel).
 */
function exportTableAsExcel(title, headers, rows) {
    var escapeHtml = function (s) {
        return String(s)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    };
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
        'xmlns="http://www.w3.org/TR/REC-html40">' +
        "<head><meta charset=\"UTF-8\"><!--[if gte mso 9]><xml>" +
        "<x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>" +
        "<x:Name>Sheet1</x:Name>" +
        "<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>" +
        "</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook>" +
        "</xml><![endif]--></head><body><table border=\"1\"><thead><tr>";
    for (var h = 0; h < headers.length; h++) {
        html += "<th>" + escapeHtml(headers[h]) + "</th>";
    }
    html += "</tr></thead><tbody>";
    for (var r = 0; r < rows.length; r++) {
        html += "<tr>";
        var row = rows[r];
        for (var c = 0; c < headers.length; c++) {
            html += "<td>" + escapeHtml(row[c] == null ? "" : String(row[c])) + "</td>";
        }
        html += "</tr>";
    }
    html += "</tbody></table></body></html>";
    var blob = new Blob([html], {
        type: "application/vnd.ms-excel;charset=utf-8;",
    });
    var safe = title.replace(/[^\w\-]+/g, "_").replace(/^_+|_+$/g, "") || "export";
    triggerDownload(blob, safe + "-" + stamp() + ".xls");
}


/***/ }),

/***/ 942:
/*!**************************************************!*\
  !*** ./lib/shared/services/permissionService.js ***!
  \**************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   accessScopeBadgeLabel: () => (/* binding */ accessScopeBadgeLabel),
/* harmony export */   getActivePermissionByEmail: () => (/* binding */ getActivePermissionByEmail),
/* harmony export */   siteAdminPermissionProfile: () => (/* binding */ siteAdminPermissionProfile)
/* harmony export */ });
/* unused harmony exports normalizePermissionRoleType, resolveCustomerRole, roleLabelFor, normalizeAccessScopeValue */
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
/* harmony import */ var _sharePointListService__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./sharePointListService */ 161);
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (undefined && undefined.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};


var permissionFields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("permissions");
var companyFields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("company");
function emailsMatch(a, b) {
    var left = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.normalizeSharePointUserEmail)(a);
    var right = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.normalizeSharePointUserEmail)(b);
    if (!left || !right)
        return false;
    return left === right;
}
/**
 * SharePoint RoleType choices today: Training Manager | Supervisor
 * (Admin / Candidate may be added later — already mapped).
 *
 * Portal routing:
 * - Admin bucket: Admin + Training Manager (legacy PAVE staff)
 * - Customer bucket: Training Manager + Supervisor + Candidate
 */
function normalizePermissionRoleType(value) {
    var role = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(value);
    if (!role)
        return null;
    var normalized = role.toLowerCase().trim();
    if (normalized === "admin" ||
        normalized === "training manager" ||
        normalized === "trainingmanager" ||
        normalized.indexOf("training manager") >= 0) {
        return "Admin";
    }
    if (normalized === "customer" ||
        normalized === "supervisor" ||
        normalized.indexOf("supervisor") >= 0 ||
        normalized === "candidate" ||
        normalized.indexOf("candidate") >= 0) {
        return "Customer";
    }
    return null;
}
function resolveCustomerRole(sharePointRole, accessScope) {
    var role = sharePointRole.toLowerCase().trim();
    var scope = accessScope.toLowerCase().trim();
    if (role === "admin")
        return null;
    if (role === "training manager" ||
        role === "trainingmanager" ||
        role.indexOf("training manager") >= 0) {
        return "TrainingManager";
    }
    if (role === "candidate" || role.indexOf("candidate") >= 0) {
        return "Candidate";
    }
    if (role === "supervisor" || role.indexOf("supervisor") >= 0) {
        if (scope.indexOf("candidate") >= 0)
            return "Candidate";
        return "Supervisor";
    }
    if (role === "customer") {
        if (scope.indexOf("candidate") >= 0)
            return "Candidate";
        if (scope.indexOf("department") >= 0)
            return "Supervisor";
        return "Supervisor";
    }
    return null;
}
function roleLabelFor(sharePointRole, customerRole) {
    if (customerRole === "TrainingManager")
        return "Training Manager";
    if (customerRole === "Supervisor")
        return "Supervisor";
    if (customerRole === "Candidate")
        return "Candidate";
    var raw = sharePointRole.trim();
    if (raw)
        return raw;
    return "Admin";
}
function normalizeAccessScopeValue(value, customerRole, isAdminOnly) {
    if (isAdminOnly)
        return "All";
    var s = value.toLowerCase().trim();
    if (s === "all" || s.indexOf("all compan") >= 0)
        return "All";
    if (s.indexOf("candidate") >= 0)
        return "CandidateOnly";
    if (s.indexOf("assigned") >= 0)
        return "AssignedCandidates";
    if (s.indexOf("department") >= 0)
        return "Department";
    if (s.indexOf("full company") >= 0 || s === "company")
        return "Company";
    if (customerRole === "Candidate")
        return "CandidateOnly";
    if (customerRole === "Supervisor")
        return "Department";
    if (customerRole === "TrainingManager")
        return "Company";
    return "Company";
}
function parseMultiChoice(value) {
    if (!value)
        return [];
    if (Array.isArray(value)) {
        return value
            .map(function (entry) {
            if (typeof entry === "string")
                return entry.trim();
            if (entry && typeof entry === "object") {
                return ((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(entry.LookupValue) ||
                    (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(entry.Title) ||
                    (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(entry.Name) ||
                    "");
            }
            return "";
        })
            .filter(Boolean);
    }
    if (typeof value === "string") {
        return value
            .split(/;|#/)
            .map(function (part) { return part.trim(); })
            .filter(function (part) { return part && !/^\d+$/.test(part); });
    }
    return [];
}
function resolveCompanyId(fields) {
    var lookupId = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields[permissionFields.companyLookupId]);
    if (lookupId)
        return lookupId;
    var companyIdRest = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields.CompanyId);
    if (companyIdRest)
        return companyIdRest;
    var companyValue = fields[permissionFields.company];
    if (companyValue &&
        typeof companyValue === "object" &&
        "LookupId" in companyValue) {
        var nestedId = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(companyValue.LookupId);
        if (nestedId)
            return nestedId;
    }
    if (typeof companyValue === "number") {
        return String(companyValue);
    }
    return (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(companyValue);
}
function resolveCompanyDisplayName(fields) {
    var _a, _b;
    var companyValue = fields[permissionFields.company];
    if (typeof companyValue === "string") {
        if (/^\d+$/.test(companyValue.trim()))
            return undefined;
        return companyValue.trim() || undefined;
    }
    if (companyValue &&
        typeof companyValue === "object" &&
        "LookupValue" in companyValue) {
        return ((_a = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(companyValue.LookupValue)) !== null && _a !== void 0 ? _a : undefined);
    }
    if (companyValue &&
        typeof companyValue === "object" &&
        "CompanyName" in companyValue) {
        return ((_b = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(companyValue.CompanyName)) !== null && _b !== void 0 ? _b : undefined);
    }
    var raw = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields.Company);
    if (raw && !/^\d+$/.test(raw))
        return raw;
    return undefined;
}
function lookupCompanyName(client, companyId) {
    return __awaiter(this, void 0, void 0, function () {
        var item, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!companyId || companyId === "0")
                        return [2 /*return*/, undefined];
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItem)(client, "company", companyId)];
                case 2:
                    item = _b.sent();
                    if (!item)
                        return [2 /*return*/, undefined];
                    return [2 /*return*/, ((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(item.fields[companyFields.companyName]) ||
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(item.fields.Title) ||
                            undefined)];
                case 3:
                    _a = _b.sent();
                    return [2 /*return*/, undefined];
                case 4: return [2 /*return*/];
            }
        });
    });
}
function mapPermissionItem(id, fields) {
    var _a, _b, _c, _d, _e, _f;
    var userEmail = (_a = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields[permissionFields.userEmail])) === null || _a === void 0 ? void 0 : _a.trim();
    var status = (_b = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields[permissionFields.status])) === null || _b === void 0 ? void 0 : _b.trim();
    var sharePointRoleType = ((_c = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields[permissionFields.roleType])) === null || _c === void 0 ? void 0 : _c.trim()) || "";
    var roleType = normalizePermissionRoleType(sharePointRoleType);
    var companyId = resolveCompanyId(fields) || "0";
    var accessScope = (_e = (_d = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields[permissionFields.accessScope])) === null || _d === void 0 ? void 0 : _d.trim()) !== null && _e !== void 0 ? _e : "Full Company";
    if (!userEmail || !status || !roleType) {
        return null;
    }
    var customerRole = resolveCustomerRole(sharePointRoleType, accessScope);
    var isAdminOnly = roleType === "Admin" && customerRole === null;
    var canAccessAdmin = roleType === "Admin" || customerRole === "TrainingManager";
    var canAccessCustomer = customerRole !== null;
    var departmentScopes = Array.from(new Set(__spreadArray(__spreadArray([], parseMultiChoice(fields[permissionFields.departments]), true), parseMultiChoice(fields[permissionFields.departmentsAllowed]), true)));
    var candidateScopeName = ((_f = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(fields[permissionFields.name])) === null || _f === void 0 ? void 0 : _f.trim()) || null;
    return {
        id: id,
        permissionItemId: id,
        userEmail: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.normalizeSharePointUserEmail)(userEmail),
        status: status,
        roleType: roleType,
        sharePointRoleType: sharePointRoleType,
        customerRole: customerRole,
        roleLabel: roleLabelFor(sharePointRoleType, customerRole),
        companyId: companyId,
        companyDisplayName: resolveCompanyDisplayName(fields),
        accessScope: accessScope,
        normalizedAccessScope: normalizeAccessScopeValue(accessScope, customerRole, isAdminOnly),
        departmentScopes: departmentScopes,
        candidateScopeName: candidateScopeName,
        canView: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asBoolean)(fields[permissionFields.canView]),
        canDownload: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asBoolean)(fields[permissionFields.canDownload]),
        canEdit: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asBoolean)(fields[permissionFields.canEdit]),
        canAccessAdmin: canAccessAdmin,
        canAccessCustomer: canAccessCustomer,
    };
}
/**
 * Active Permissions List row for signed-in email.
 * Resolves Company Name from Company List when REST only returns CompanyId.
 */
function getActivePermissionByEmail(client, email) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var normalizedEmail, items, matches, i, permission, chosen, name_1;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    normalizedEmail = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.normalizeSharePointUserEmail)(email);
                    if (!normalizedEmail)
                        return [2 /*return*/, null];
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "permissions", {
                            filter: [
                                (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(permissionFields.userEmail, normalizedEmail),
                                (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(permissionFields.status, "Active"),
                            ].join(" and "),
                            top: 100,
                            maxItems: 100,
                        })];
                case 1:
                    items = _b.sent();
                    if (!(items.length === 0)) return [3 /*break*/, 3];
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "permissions", {
                            filter: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(permissionFields.status, "Active"),
                            top: 5000,
                            maxItems: 5000,
                        })];
                case 2:
                    items = _b.sent();
                    _b.label = 3;
                case 3:
                    matches = [];
                    for (i = 0; i < items.length; i++) {
                        permission = mapPermissionItem(items[i].id, items[i].fields);
                        if (permission &&
                            emailsMatch(permission.userEmail, normalizedEmail) &&
                            permission.status.toLowerCase() === "active") {
                            matches.push(permission);
                        }
                    }
                    if (matches.length === 0)
                        return [2 /*return*/, null];
                    chosen = (_a = matches.find(function (p) { return p.roleType === "Admin"; })) !== null && _a !== void 0 ? _a : matches[0];
                    if (!(!chosen.companyDisplayName &&
                        chosen.companyId &&
                        chosen.companyId !== "0")) return [3 /*break*/, 5];
                    return [4 /*yield*/, lookupCompanyName(client, chosen.companyId)];
                case 4:
                    name_1 = _b.sent();
                    if (name_1) {
                        return [2 /*return*/, __assign(__assign({}, chosen), { companyDisplayName: name_1 })];
                    }
                    _b.label = 5;
                case 5: return [2 /*return*/, chosen];
            }
        });
    });
}
function siteAdminPermissionProfile(email) {
    return {
        id: "site-admin",
        permissionItemId: "site-admin",
        userEmail: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.normalizeSharePointUserEmail)(email),
        status: "Active",
        roleType: "Admin",
        sharePointRoleType: "Admin",
        customerRole: null,
        roleLabel: "Admin",
        companyId: "0",
        companyDisplayName: "All companies",
        accessScope: "All",
        normalizedAccessScope: "All",
        departmentScopes: [],
        candidateScopeName: null,
        canView: true,
        canDownload: true,
        canEdit: true,
        canAccessAdmin: true,
        canAccessCustomer: false,
    };
}
/** Access badge text for customer header. */
function accessScopeBadgeLabel(permission) {
    var scope = permission.normalizedAccessScope;
    if (scope === "Company" || scope === "All")
        return "Company-wide";
    if (scope === "Department") {
        if (permission.departmentScopes.length > 0) {
            return permission.departmentScopes.join(", ") + " department";
        }
        return "Department";
    }
    if (scope === "AssignedCandidates")
        return "Assigned candidates";
    if (scope === "CandidateOnly")
        return "Own records only";
    return permission.accessScope || "Company";
}


/***/ }),

/***/ 350:
/*!**************************************************!*\
  !*** ./lib/shared/services/portalDataService.js ***!
  \**************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getAdminSchemaColumns: () => (/* binding */ getAdminSchemaColumns),
/* harmony export */   loadCustomerDashboardData: () => (/* binding */ loadCustomerDashboardData),
/* harmony export */   loadDashboardCounts: () => (/* binding */ loadDashboardCounts),
/* harmony export */   loadPortalListRows: () => (/* binding */ loadPortalListRows)
/* harmony export */ });
/* unused harmony export listSchemaFieldNames */
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
/* harmony import */ var _sharePointListService__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./sharePointListService */ 161);
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};


function cell(value) {
    var _a;
    return (_a = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asString)(value)) !== null && _a !== void 0 ? _a : "";
}
function readField(fields, internalName) {
    if (fields[internalName] !== undefined && fields[internalName] !== null) {
        return fields[internalName];
    }
    // Lookup id companion (Company → CompanyId)
    var idKey = internalName + "Id";
    if (fields[idKey] !== undefined && fields[idKey] !== null) {
        return fields[idKey];
    }
    return fields[internalName];
}
/**
 * Lookup expands — only use when paired with a matching $select.
 * Prefer no expand by default (CompanyId etc. still return without it).
 */
var EXPAND_BY_LIST = {};
/**
 * Returns every mapped schema field for Admin tables (closer to SharePoint list view).
 */
function getAdminSchemaColumns(listKey) {
    var list = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointList)(listKey);
    var fields = list.fields;
    var columns = [];
    var headers = [];
    for (var key in fields) {
        if (!Object.prototype.hasOwnProperty.call(fields, key))
            continue;
        if (key === "id" || key === "companyLookupId")
            continue;
        var internal = fields[key];
        if (columns.indexOf(internal) >= 0)
            continue;
        columns.push(internal);
        headers.push((0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFieldLabel)(listKey, key));
    }
    return { columns: columns, headers: headers };
}
/**
 * Loads list rows for portal tables (admin: all pages; customer: company + visible).
 */
function loadPortalListRows(client, listKey, options) {
    return __awaiter(this, void 0, void 0, function () {
        var filters, expand, loadAll, items;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    filters = [];
                    if (options.companyId && options.companyIdFieldInternalName) {
                        filters.push(options.companyIdFieldInternalName + " eq " + Number(options.companyId));
                    }
                    else if (options.companyName && options.companyFieldInternalName) {
                        filters.push((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(options.companyFieldInternalName, options.companyName));
                    }
                    if (options.customerVisibleOnly && options.visibleFieldInternalName) {
                        filters.push(options.visibleFieldInternalName + " eq 1");
                    }
                    expand = EXPAND_BY_LIST[listKey] || [];
                    loadAll = options.loadAll !== false && !options.companyName && !options.companyId;
                    return [4 /*yield*/, (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, listKey, {
                            filter: filters.length ? filters.join(" and ") : undefined,
                            top: options.top != null ? options.top : loadAll ? 5000 : 500,
                            maxItems: loadAll ? 20000 : options.top != null ? options.top : 500,
                            expand: expand.length ? expand : undefined,
                        })];
                case 1:
                    items = _a.sent();
                    return [2 /*return*/, items.map(function (item) { return ({
                            id: item.id,
                            cells: options.columns.map(function (col) { return cell(readField(item.fields, col)); }),
                            fields: item.fields,
                        }); })];
            }
        });
    });
}
function loadDashboardCounts(client, companyName) {
    return __awaiter(this, void 0, void 0, function () {
        var workforce, matrix, events, offers, docs, companyFilter, adminTop, adminMax, _a, w, m, e, o, d;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    workforce = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("workforce");
                    matrix = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("trainingMatrix");
                    events = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("events");
                    offers = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("offersPromotions");
                    docs = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("customerDocuments");
                    companyFilter = function (field) {
                        return companyName ? (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(field, companyName) : undefined;
                    };
                    adminTop = companyName ? 5000 : 5000;
                    adminMax = companyName ? 5000 : 20000;
                    return [4 /*yield*/, Promise.all([
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "workforce", {
                                filter: companyFilter(workforce.companyName),
                                top: adminTop,
                                maxItems: adminMax,
                            }),
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "trainingMatrix", {
                                filter: companyFilter(matrix.companyName),
                                top: adminTop,
                                maxItems: adminMax,
                            }),
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "events", {
                                filter: companyName
                                    ? [
                                        (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(events.eventCompany, companyName),
                                        events.customerVisible + " eq 1",
                                    ].join(" and ")
                                    : undefined,
                                top: adminTop,
                                maxItems: adminMax,
                            }),
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "offersPromotions", {
                                filter: companyName ? offers.customerVisible + " eq 1" : undefined,
                                top: 2000,
                                maxItems: 2000,
                            }).catch(function () { return []; }),
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "customerDocuments", {
                                filter: companyName
                                    ? [
                                        (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(docs.company, companyName),
                                        docs.customerVisible + " eq 1",
                                    ].join(" and ")
                                    : undefined,
                                top: adminTop,
                                maxItems: adminMax,
                            }),
                        ])];
                case 1:
                    _a = _b.sent(), w = _a[0], m = _a[1], e = _a[2], o = _a[3], d = _a[4];
                    return [2 /*return*/, {
                            workforce: w.length,
                            matrix: m.length,
                            events: e.length,
                            offers: o.filter(function (x) {
                                return companyName ? (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asBoolean)(x.fields[offers.customerVisible]) : true;
                            }).length,
                            documents: d.length,
                        }];
            }
        });
    });
}
function daysUntil(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime()))
        return null;
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.ceil((d.getTime() - today.getTime()) / 86400000);
}
function nvqProgressPct(stage) {
    var s = stage.toLowerCase();
    if (s.indexOf("complete") >= 0 || s.indexOf("achieved") >= 0)
        return 100;
    if (s.indexOf("assess") >= 0 || s.indexOf("portfolio") >= 0)
        return 75;
    if (s.indexOf("induct") >= 0 || s.indexOf("register") >= 0)
        return 25;
    if (s.indexOf("progress") >= 0 || s.indexOf("active") >= 0)
        return 50;
    if (!s)
        return 0;
    return 40;
}
/**
 * Customer dashboard payload — company-scoped SharePoint lists (real data).
 */
function loadCustomerDashboardData(client, companyName, companyId) {
    return __awaiter(this, void 0, void 0, function () {
        var matrix, docs, nvq, events, offers, workforce, safe, lookupId, docsFilter, eventsFilter, _a, wItems, mItems, dItems, nItems, eItems, oItems, expiringSoon, missingData, matrixRows, documentTiles, nvqRows, eventRows, offerCards;
        var _this = this;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    matrix = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("trainingMatrix");
                    docs = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("customerDocuments");
                    nvq = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("nvqRegister");
                    events = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("events");
                    offers = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("offersPromotions");
                    workforce = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFields)("workforce");
                    safe = function (p, fallback) { return __awaiter(_this, void 0, void 0, function () {
                        var _a;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    _b.trys.push([0, 2, , 3]);
                                    return [4 /*yield*/, p];
                                case 1: return [2 /*return*/, _b.sent()];
                                case 2:
                                    _a = _b.sent();
                                    return [2 /*return*/, fallback];
                                case 3: return [2 /*return*/];
                            }
                        });
                    }); };
                    lookupId = companyId && companyId !== "0" && /^\d+$/.test(companyId)
                        ? Number(companyId)
                        : NaN;
                    docsFilter = !isNaN(lookupId)
                        ? ["CompanyId eq " + lookupId, docs.customerVisible + " eq 1"].join(" and ")
                        : [
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(docs.company, companyName),
                            docs.customerVisible + " eq 1",
                        ].join(" and ");
                    eventsFilter = !isNaN(lookupId)
                        ? [
                            "EventCompanyId eq " + lookupId,
                            events.customerVisible + " eq 1",
                        ].join(" and ")
                        : [
                            (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(events.eventCompany, companyName),
                            events.customerVisible + " eq 1",
                        ].join(" and ");
                    return [4 /*yield*/, Promise.all([
                            safe((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "workforce", {
                                filter: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(workforce.companyName, companyName),
                                top: 5000,
                                maxItems: 5000,
                            }), []),
                            safe((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "trainingMatrix", {
                                filter: (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(matrix.companyName, companyName),
                                top: 5000,
                                maxItems: 5000,
                            }), []),
                            safe((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "customerDocuments", {
                                filter: docsFilter,
                                top: 200,
                                maxItems: 200,
                            }), []),
                            safe((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "nvqRegister", {
                                filter: [
                                    (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.fieldEqualsFilter)(nvq.companyName, companyName),
                                    nvq.customerVisible + " eq 1",
                                ].join(" and "),
                                top: 200,
                                maxItems: 200,
                            }), []),
                            safe((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "events", {
                                filter: eventsFilter,
                                top: 200,
                                maxItems: 200,
                            }), []),
                            safe((0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.getListItems)(client, "offersPromotions", {
                                filter: offers.customerVisible + " eq 1",
                                top: 50,
                                maxItems: 50,
                            }), []),
                        ])];
                case 1:
                    _a = _b.sent(), wItems = _a[0], mItems = _a[1], dItems = _a[2], nItems = _a[3], eItems = _a[4], oItems = _a[5];
                    expiringSoon = 0;
                    missingData = 0;
                    matrixRows = mItems.map(function (item) {
                        var status = cell(item.fields[matrix.overallStatus]);
                        var expiry = cell(item.fields[matrix.nextExpiryDate]);
                        var needsReview = (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asBoolean)(item.fields[matrix.needsReview]);
                        var days = expiry ? daysUntil(expiry) : null;
                        if (needsReview || /missing|review/i.test(status)) {
                            missingData += 1;
                        }
                        else if (days !== null && days >= 0 && days <= 60) {
                            expiringSoon += 1;
                        }
                        else if (days !== null && days < 0) {
                            missingData += 1;
                        }
                        return {
                            id: item.id,
                            cells: [
                                cell(item.fields[matrix.candidateName]),
                                status || (needsReview ? "Missing Data" : "Compliant"),
                                expiry,
                            ],
                            fields: item.fields,
                        };
                    });
                    documentTiles = dItems
                        .slice(0, 8)
                        .map(function (item) {
                        var name = cell(item.fields[docs.fileLeafRef]) ||
                            cell(item.fields[docs.title]) ||
                            "Document";
                        var type = cell(item.fields[docs.documentType]) || "";
                        var fsObj = item.fields[docs.fsObjType];
                        var isFolder = fsObj === 1 || fsObj === "1";
                        var isPdf = !isFolder && (/\.pdf$/i.test(name) || /pdf/i.test(type));
                        return {
                            id: item.id,
                            label: name,
                            meta: type || (isFolder ? "Folder" : isPdf ? "PDF" : "File"),
                            kind: isPdf ? "pdf" : "folder",
                        };
                    });
                    nvqRows = nItems.slice(0, 5).map(function (item) {
                        var stage = cell(item.fields[nvq.stageOfNvq]);
                        var title = cell(item.fields[nvq.nvqTitle]);
                        return {
                            id: item.id,
                            name: cell(item.fields[nvq.candidateName]) || "—",
                            course: [title, stage].filter(Boolean).join(" · ") || "NVQ",
                            pct: nvqProgressPct(stage),
                        };
                    });
                    eventRows = eItems.slice(0, 5).map(function (item) {
                        var dateRaw = cell(item.fields[events.eventDate]);
                        var when = dateRaw;
                        var d = new Date(dateRaw);
                        if (!isNaN(d.getTime())) {
                            when = d.toLocaleString("en-GB", {
                                day: "numeric",
                                month: "short",
                                year: "numeric",
                                hour: "2-digit",
                                minute: "2-digit",
                            });
                        }
                        return {
                            id: item.id,
                            title: cell(item.fields[events.title]) || "Event",
                            when: when,
                            where: cell(item.fields[events.trainingAddress]) || "—",
                            dateRaw: dateRaw,
                        };
                    });
                    offerCards = oItems
                        .filter(function (x) { return (0,_sharePointListService__WEBPACK_IMPORTED_MODULE_1__.asBoolean)(x.fields[offers.customerVisible]); })
                        .slice(0, 4)
                        .map(function (item) {
                        var category = cell(item.fields[offers.category]) || "Offer";
                        var title = cell(item.fields[offers.title]) ||
                            cell(item.fields[offers.shortDescription]) ||
                            "Promotion";
                        return {
                            id: item.id,
                            badge: category.toUpperCase(),
                            title: title,
                            code: "Learn more →",
                        };
                    });
                    return [2 /*return*/, {
                            counts: {
                                workforce: wItems.length,
                                matrix: mItems.length,
                                events: eItems.length,
                                offers: offerCards.length,
                                documents: dItems.length,
                                expiringSoon: expiringSoon,
                                missingData: missingData,
                            },
                            matrixRows: matrixRows,
                            documentTiles: documentTiles,
                            nvqRows: nvqRows,
                            eventRows: eventRows,
                            offerCards: offerCards,
                        }];
            }
        });
    });
}
/** Convenience: schema field names for a list (debug / future edit forms). */
function listSchemaFieldNames(listKey) {
    return (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_0__.getSharePointFieldInternalNames)(listKey);
}


/***/ }),

/***/ 161:
/*!******************************************************!*\
  !*** ./lib/shared/services/sharePointListService.js ***!
  \******************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   asBoolean: () => (/* binding */ asBoolean),
/* harmony export */   asString: () => (/* binding */ asString),
/* harmony export */   createListItem: () => (/* binding */ createListItem),
/* harmony export */   deleteListItem: () => (/* binding */ deleteListItem),
/* harmony export */   fieldEqualsFilter: () => (/* binding */ fieldEqualsFilter),
/* harmony export */   getListItem: () => (/* binding */ getListItem),
/* harmony export */   getListItems: () => (/* binding */ getListItems),
/* harmony export */   normalizeSharePointUserEmail: () => (/* binding */ normalizeSharePointUserEmail),
/* harmony export */   updateListItem: () => (/* binding */ updateListItem)
/* harmony export */ });
/* unused harmony export sanitizeWriteFields */
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-http */ 272);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};


function escapeOData(value) {
    return value.replace(/'/g, "''");
}
/**
 * Build `/_api/web/lists/...` root. Prefer GUID when set (titles with "/"
 * like "Offers / Promotions" 404 if used raw in getbytitle).
 */
function listApiRoot(client, listKey) {
    var list = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey];
    if (list.listId) {
        return (client.webUrl + "/_api/web/lists(guid'" + list.listId + "')");
    }
    // Encode so spaces and "/" don't break the path
    return (client.webUrl +
        "/_api/web/lists/getbytitle('" +
        encodeURIComponent(escapeOData(list.listName)) +
        "')");
}
function restHeaders(extra) {
    var headers = {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "odata-version": "",
    };
    if (extra) {
        for (var key in extra) {
            if (Object.prototype.hasOwnProperty.call(extra, key)) {
                headers[key] = extra[key];
            }
        }
    }
    return headers;
}
function writeHeaders(extra) {
    // Use plain application/json for writes — SPO can throw
    // JsonReaderException when Content-Type includes odata=nometadata.
    var headers = {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json",
        "odata-version": "",
    };
    if (extra) {
        for (var key in extra) {
            if (Object.prototype.hasOwnProperty.call(extra, key)) {
                headers[key] = extra[key];
            }
        }
    }
    return headers;
}
/**
 * Fields that must not be written as plain strings via REST (complex types).
 */
var NON_WRITABLE_FIELDS = {
    CompanyLogo: true,
    Attachments: true,
    FileLeafRef: true,
    FileRef: true,
    FileDirRef: true,
    AuthorId: true,
    EditorId: true,
    ContentTypeId: true,
    ContentType: true,
    GUID: true,
    UniqueId: true,
    Created: true,
    Modified: true,
    ID: true,
    Id: true,
    CompanyLookupId: true,
};
/**
 * Drop empty / complex values and ensure Title is present for list creates.
 */
function sanitizeWriteFields(fields, options) {
    var clean = {};
    for (var key in fields) {
        if (!Object.prototype.hasOwnProperty.call(fields, key))
            continue;
        // CompanyLogo is writable only as a stringified thumbnail JSON blob
        if (key === "CompanyLogo") {
            var v = fields[key];
            if (typeof v === "string" && v.trim().charAt(0) === "{") {
                clean[key] = v;
            }
            continue;
        }
        if (NON_WRITABLE_FIELDS[key])
            continue;
        if (key.indexOf(".") >= 0)
            continue;
        if (key.indexOf("@odata.") === 0)
            continue;
        if (key === "__metadata")
            continue;
        var value = fields[key];
        if (value === undefined)
            continue;
        if (value === null) {
            if (!options || !options.forCreate) {
                clean[key] = null;
            }
            continue;
        }
        if (typeof value === "string" && value.trim() === "")
            continue;
        if (typeof value === "object" && !Array.isArray(value)) {
            // Skip nested objects (lookups/thumbnails) — use *Id fields instead
            continue;
        }
        clean[key] = value;
    }
    if (options &&
        options.forCreate &&
        (clean.Title === undefined || clean.Title === null || clean.Title === "")) {
        var name_1 = clean.CompanyName || clean.CandidateName || clean.UserEmail;
        if (typeof name_1 === "string" && name_1.trim()) {
            clean.Title = name_1.trim();
        }
    }
    return clean;
}
function mapItem(item, fallbackId) {
    var id = String(item.Id != null
        ? item.Id
        : item.ID != null
            ? item.ID
            : fallbackId != null
                ? fallbackId
                : "");
    var fields = {};
    for (var key in item) {
        if (Object.prototype.hasOwnProperty.call(item, key) &&
            key !== "@odata.etag") {
            fields[key] = item[key];
        }
    }
    return { id: id, fields: fields };
}
function getJson(client, url) {
    return __awaiter(this, void 0, void 0, function () {
        var response, text;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, client.spHttpClient.get(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, { headers: restHeaders() })];
                case 1:
                    response = _a.sent();
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    text = _a.sent();
                    throw new Error("SharePoint REST " + response.status + ": " + text);
                case 3: return [4 /*yield*/, response.json()];
                case 4: return [2 /*return*/, (_a.sent())];
            }
        });
    });
}
/**
 * Reads list items via SharePoint REST (current user context).
 * Follows nextLink pages so Admin gets full list data, not only the first page.
 */
function getListItems(client, listKey, options) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, top, maxItems, params, url, rows, json, e_1, message, items, i, next;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    top = options && options.top ? options.top : 5000;
                    maxItems = options && options.maxItems ? options.maxItems : 20000;
                    params = ["$top=" + top];
                    if (options && options.filter) {
                        params.push("$filter=" + encodeURIComponent(options.filter));
                    }
                    if (options && options.select && options.select.length > 0) {
                        params.push("$select=" + options.select.map(encodeURIComponent).join(","));
                    }
                    if (options && options.expand && options.expand.length > 0) {
                        params.push("$expand=" + options.expand.map(encodeURIComponent).join(","));
                    }
                    url = listApiRoot(client, listKey) + "/items?" + params.join("&");
                    rows = [];
                    _a.label = 1;
                case 1:
                    if (!(url && rows.length < maxItems)) return [3 /*break*/, 6];
                    json = void 0;
                    _a.label = 2;
                case 2:
                    _a.trys.push([2, 4, , 5]);
                    return [4 /*yield*/, getJson(client, url)];
                case 3:
                    json = _a.sent();
                    return [3 /*break*/, 5];
                case 4:
                    e_1 = _a.sent();
                    message = e_1 instanceof Error ? e_1.message : String(e_1);
                    throw new Error('Failed to read "' + listName + '": ' + message);
                case 5:
                    items = json.value || [];
                    for (i = 0; i < items.length; i++) {
                        rows.push(mapItem(items[i]));
                        if (rows.length >= maxItems) {
                            break;
                        }
                    }
                    next = json["@odata.nextLink"] ||
                        json["odata.nextLink"];
                    url = next || "";
                    return [3 /*break*/, 1];
                case 6: return [2 /*return*/, rows];
            }
        });
    });
}
function getListItem(client, listKey, itemId) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, url, response, text, item;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    url = listApiRoot(client, listKey) +
                        "/items(" +
                        encodeURIComponent(itemId) +
                        ")";
                    return [4 /*yield*/, client.spHttpClient.get(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, { headers: restHeaders() })];
                case 1:
                    response = _a.sent();
                    if (response.status === 404) {
                        return [2 /*return*/, null];
                    }
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    text = _a.sent();
                    throw new Error("Failed to read item " +
                        itemId +
                        ' from "' +
                        listName +
                        '" (' +
                        response.status +
                        "): " +
                        text);
                case 3: return [4 /*yield*/, response.json()];
                case 4:
                    item = (_a.sent());
                    return [2 /*return*/, mapItem(item, itemId)];
            }
        });
    });
}
function createListItem(client, listKey, fields) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, clean, bodyText, url, response, text, item;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    clean = sanitizeWriteFields(fields, { forCreate: true });
                    if (Object.keys(clean).length === 0) {
                        throw new Error('Nothing to create in "' +
                            listName +
                            '". Fill at least Title or Company Name / Candidate Name.');
                    }
                    bodyText = JSON.stringify(clean);
                    if (!bodyText || bodyText.charAt(0) !== "{") {
                        throw new Error("Invalid create payload (not a JSON object).");
                    }
                    url = listApiRoot(client, listKey) + "/items";
                    return [4 /*yield*/, client.spHttpClient.post(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, {
                            headers: writeHeaders(),
                            body: bodyText,
                        })];
                case 1:
                    response = _a.sent();
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    text = _a.sent();
                    throw new Error('Failed to create item in "' +
                        listName +
                        '" (' +
                        response.status +
                        "): " +
                        text);
                case 3: return [4 /*yield*/, response.json()];
                case 4:
                    item = (_a.sent());
                    return [2 /*return*/, {
                            id: String(item.Id != null ? item.Id : item.ID != null ? item.ID : ""),
                            fields: clean,
                        }];
            }
        });
    });
}
function updateListItem(client, listKey, itemId, fields) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, clean, bodyText, url, response, text;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    clean = sanitizeWriteFields(fields, { forCreate: false });
                    if (Object.keys(clean).length === 0) {
                        throw new Error("Nothing to update — no writable field values provided.");
                    }
                    bodyText = JSON.stringify(clean);
                    if (!bodyText || bodyText.charAt(0) !== "{") {
                        throw new Error("Invalid update payload (not a JSON object).");
                    }
                    url = listApiRoot(client, listKey) +
                        "/items(" +
                        encodeURIComponent(itemId) +
                        ")";
                    return [4 /*yield*/, client.spHttpClient.post(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, {
                            headers: writeHeaders({
                                "IF-MATCH": "*",
                                "X-HTTP-Method": "MERGE",
                            }),
                            body: bodyText,
                        })];
                case 1:
                    response = _a.sent();
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    text = _a.sent();
                    throw new Error("Failed to update item " +
                        itemId +
                        ' in "' +
                        listName +
                        '" (' +
                        response.status +
                        "): " +
                        text);
                case 3: return [2 /*return*/];
            }
        });
    });
}
/**
 * Deletes a list item via SharePoint REST DELETE.
 */
function deleteListItem(client, listKey, itemId) {
    return __awaiter(this, void 0, void 0, function () {
        var listName, url, response, text;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    listName = _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.SHAREPOINT_LISTS[listKey].listName;
                    url = listApiRoot(client, listKey) +
                        "/items(" +
                        encodeURIComponent(itemId) +
                        ")";
                    return [4 /*yield*/, client.spHttpClient.post(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, {
                            headers: restHeaders({
                                "IF-MATCH": "*",
                                "X-HTTP-Method": "DELETE",
                            }),
                        })];
                case 1:
                    response = _a.sent();
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    text = _a.sent();
                    throw new Error("Failed to delete item " +
                        itemId +
                        ' in "' +
                        listName +
                        '" (' +
                        response.status +
                        "): " +
                        text);
                case 3: return [2 /*return*/];
            }
        });
    });
}
function asString(value) {
    if (value === null || value === undefined)
        return null;
    if (typeof value === "string")
        return value.trim() || null;
    if (typeof value === "number" || typeof value === "boolean") {
        return String(value);
    }
    if (Array.isArray(value)) {
        var parts = [];
        for (var i = 0; i < value.length; i++) {
            var part = asString(value[i]);
            if (part)
                parts.push(part);
        }
        return parts.length ? parts.join(", ") : null;
    }
    if (typeof value === "object" && value !== null) {
        var record = value;
        if (typeof record.Title === "string")
            return record.Title.trim() || null;
        if (typeof record.LookupValue === "string") {
            return record.LookupValue.trim() || null;
        }
        if (typeof record.CompanyName === "string") {
            return record.CompanyName.trim() || null;
        }
        if (typeof record.Email === "string")
            return record.Email.trim() || null;
    }
    return null;
}
function asBoolean(value) {
    if (typeof value === "boolean")
        return value;
    if (typeof value === "number")
        return value === 1;
    if (typeof value === "string") {
        var n = value.trim().toLowerCase();
        return n === "true" || n === "yes" || n === "1";
    }
    return false;
}
function fieldEqualsFilter(internalName, value) {
    return internalName + " eq '" + escapeOData(value) + "'";
}
/**
 * Turns SharePoint loginName / UPN into a plain email when possible.
 */
function normalizeSharePointUserEmail(raw) {
    var value = (raw || "").trim();
    if (!value)
        return "";
    var membership = value.match(/\|membership\|([^|]+)$/i);
    if (membership && membership[1]) {
        value = membership[1];
    }
    else {
        var pipe = value.lastIndexOf("|");
        if (pipe >= 0 && value.indexOf("@") > pipe) {
            value = value.substring(pipe + 1);
        }
    }
    var hash = value.lastIndexOf("#");
    if (hash >= 0 && value.indexOf("@") > hash) {
        value = value.substring(hash + 1);
    }
    return value.trim().toLowerCase();
}


/***/ }),

/***/ 179:
/*!*****************************************!*\
  !*** ./lib/shared/ui/AdminDataTable.js ***!
  \*****************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AdminDataTable: () => (/* binding */ AdminDataTable)
/* harmony export */ });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 650);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
/* harmony import */ var _services_companyLogoService__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../services/companyLogoService */ 827);
/* harmony import */ var _services_exportService__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../services/exportService */ 237);
/* harmony import */ var _services_sharePointListService__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../services/sharePointListService */ 161);
/* harmony import */ var _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./portal.module.scss */ 256);
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};






var READONLY_FIELDS = {
    ID: true,
    Id: true,
    FileLeafRef: true,
    FileRef: true,
    FileDirRef: true,
    UniqueId: true,
    GUID: true,
    Created: true,
    Modified: true,
    AuthorId: true,
    EditorId: true,
    ContentTypeId: true,
    Attachments: true,
    CompanyLookupId: true,
    CompanyLogo: true,
};
var BOOLEAN_HINTS = [
    "CustomerVisible",
    "Customer_x0020_Visible",
    "CanView",
    "CanDownload",
    "CanEdit",
    "Active",
    "DoNotSync",
    "ReceiveExpiryNotifications",
    "ReceiveDocumentNotifications",
    "CustomerNotificationsEnabled",
];
function isBooleanField(internalName) {
    return BOOLEAN_HINTS.indexOf(internalName) >= 0;
}
function isLookupIdField(internalName) {
    return (/Id$/i.test(internalName) &&
        internalName !== "ID" &&
        internalName !== "Id");
}
function fieldToFormValue(value) {
    if (value === null || value === undefined)
        return "";
    if (typeof value === "boolean")
        return value ? "true" : "false";
    if (typeof value === "object") {
        var s = (0,_services_sharePointListService__WEBPACK_IMPORTED_MODULE_3__.asString)(value);
        return s || "";
    }
    return String(value);
}
function buildWritePayload(columns, form, original) {
    var payload = {};
    for (var i = 0; i < columns.length; i++) {
        var name_1 = columns[i];
        if (READONLY_FIELDS[name_1])
            continue;
        if (!Object.prototype.hasOwnProperty.call(form, name_1))
            continue;
        var idCompanion = name_1 + "Id";
        if (isLookupIdField(name_1)) {
            var raw = (form[name_1] || "").trim();
            payload[name_1] =
                raw === "" ? null : /^\d+$/.test(raw) ? Number(raw) : raw;
            continue;
        }
        if (Object.prototype.hasOwnProperty.call(original, idCompanion)) {
            var raw = (form[name_1] || "").trim();
            payload[idCompanion] =
                raw === "" ? null : /^\d+$/.test(raw) ? Number(raw) : null;
            continue;
        }
        if (isBooleanField(name_1)) {
            var raw = (form[name_1] || "").trim().toLowerCase();
            if (raw === "")
                continue;
            payload[name_1] =
                raw === "true" || raw === "1" || raw === "yes" || raw === "on";
            continue;
        }
        var text = (form[name_1] || "").trim();
        if (text === "")
            continue;
        payload[name_1] = text;
    }
    return payload;
}
var AdminDataTable = function (props) {
    var client = props.client, listKey = props.listKey, title = props.title, headers = props.headers, columns = props.columns, rows = props.rows, loading = props.loading, error = props.error, onRefresh = props.onRefresh;
    var _a = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), drawerOpen = _a[0], setDrawerOpen = _a[1];
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), editingId = _b[0], setEditingId = _b[1];
    var _c = react__WEBPACK_IMPORTED_MODULE_0__.useState({}), form = _c[0], setForm = _c[1];
    var _d = react__WEBPACK_IMPORTED_MODULE_0__.useState({}), original = _d[0], setOriginal = _d[1];
    var _e = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), busy = _e[0], setBusy = _e[1];
    var _f = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), actionError = _f[0], setActionError = _f[1];
    var _g = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), actionOk = _g[0], setActionOk = _g[1];
    var _h = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), logoFile = _h[0], setLogoFile = _h[1];
    var _j = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), logoPreview = _j[0], setLogoPreview = _j[1];
    var _k = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), localPreviewUrl = _k[0], setLocalPreviewUrl = _k[1];
    var listLabel = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.getSharePointList)(listKey).displayName;
    var supportsCompanyLogo = listKey === "company";
    var clearLogoSelection = function () {
        if (localPreviewUrl) {
            URL.revokeObjectURL(localPreviewUrl);
        }
        setLocalPreviewUrl(null);
        setLogoFile(null);
    };
    var openCreate = function () {
        var blank = {};
        for (var i = 0; i < columns.length; i++) {
            if (!READONLY_FIELDS[columns[i]])
                blank[columns[i]] = "";
        }
        setEditingId(null);
        setOriginal({});
        setForm(blank);
        clearLogoSelection();
        setLogoPreview(null);
        setActionError(null);
        setActionOk(null);
        setDrawerOpen(true);
    };
    var openEdit = function (row) { return __awaiter(void 0, void 0, void 0, function () {
        var item, fields, next, i, name_2, idKey, key, e_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setBusy(true);
                    setActionError(null);
                    setActionOk(null);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, (0,_services_sharePointListService__WEBPACK_IMPORTED_MODULE_3__.getListItem)(client, listKey, row.id)];
                case 2:
                    item = _a.sent();
                    fields = item ? item.fields : row.fields || {};
                    next = {};
                    for (i = 0; i < columns.length; i++) {
                        name_2 = columns[i];
                        if (READONLY_FIELDS[name_2])
                            continue;
                        idKey = name_2 + "Id";
                        if (fields[idKey] !== undefined && fields[idKey] !== null) {
                            next[name_2] = fieldToFormValue(fields[idKey]);
                        }
                        else {
                            next[name_2] = fieldToFormValue(fields[name_2]);
                        }
                    }
                    for (key in fields) {
                        if (Object.prototype.hasOwnProperty.call(fields, key) &&
                            isLookupIdField(key) &&
                            !READONLY_FIELDS[key]) {
                            next[key] = fieldToFormValue(fields[key]);
                        }
                    }
                    setEditingId(row.id);
                    setOriginal(fields);
                    setForm(next);
                    clearLogoSelection();
                    setLogoPreview((0,_services_companyLogoService__WEBPACK_IMPORTED_MODULE_2__.thumbnailPreviewUrl)(client.webUrl, fields.CompanyLogo));
                    setDrawerOpen(true);
                    return [3 /*break*/, 5];
                case 3:
                    e_1 = _a.sent();
                    setActionError(e_1 instanceof Error ? e_1.message : "Failed to load item.");
                    return [3 /*break*/, 5];
                case 4:
                    setBusy(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var onLogoChosen = function (fileList) {
        var file = fileList && fileList.length > 0 ? fileList[0] : null;
        clearLogoSelection();
        if (!file)
            return;
        if (file.type && file.type.indexOf("image/") !== 0) {
            setActionError("Please choose an image file (PNG, JPG, WEBP, etc.).");
            return;
        }
        setActionError(null);
        setLogoFile(file);
        setLocalPreviewUrl(URL.createObjectURL(file));
    };
    var save = function () { return __awaiter(void 0, void 0, void 0, function () {
        var payload, key, raw, itemId, created, e_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setBusy(true);
                    setActionError(null);
                    setActionOk(null);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 8, 9, 10]);
                    payload = buildWritePayload(columns, form, original);
                    for (key in form) {
                        if (Object.prototype.hasOwnProperty.call(form, key) &&
                            isLookupIdField(key) &&
                            !READONLY_FIELDS[key]) {
                            raw = (form[key] || "").trim();
                            if (raw === "") {
                                if (editingId)
                                    payload[key] = null;
                            }
                            else {
                                payload[key] = /^\d+$/.test(raw) ? Number(raw) : raw;
                            }
                        }
                    }
                    itemId = editingId;
                    if (!editingId) return [3 /*break*/, 3];
                    return [4 /*yield*/, (0,_services_sharePointListService__WEBPACK_IMPORTED_MODULE_3__.updateListItem)(client, listKey, editingId, payload)];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, (0,_services_sharePointListService__WEBPACK_IMPORTED_MODULE_3__.createListItem)(client, listKey, payload)];
                case 4:
                    created = _a.sent();
                    itemId = created.id;
                    _a.label = 5;
                case 5:
                    if (!(supportsCompanyLogo && logoFile && itemId)) return [3 /*break*/, 7];
                    return [4 /*yield*/, (0,_services_companyLogoService__WEBPACK_IMPORTED_MODULE_2__.uploadAndSetListImage)(client, listKey, itemId, "CompanyLogo", logoFile)];
                case 6:
                    _a.sent();
                    _a.label = 7;
                case 7:
                    setActionOk(editingId
                        ? "Saved item #" + editingId + (logoFile ? " (logo updated)" : "")
                        : "Created item #" + itemId + (logoFile ? " (logo uploaded)" : ""));
                    setDrawerOpen(false);
                    clearLogoSelection();
                    onRefresh();
                    return [3 /*break*/, 10];
                case 8:
                    e_2 = _a.sent();
                    setActionError(e_2 instanceof Error ? e_2.message : "Save failed.");
                    return [3 /*break*/, 10];
                case 9:
                    setBusy(false);
                    return [7 /*endfinally*/];
                case 10: return [2 /*return*/];
            }
        });
    }); };
    var remove = function (row) { return __awaiter(void 0, void 0, void 0, function () {
        var ok, e_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    ok = window.confirm("Delete item #" +
                        row.id +
                        " from " +
                        listLabel +
                        "? This cannot be undone.");
                    if (!ok)
                        return [2 /*return*/];
                    setBusy(true);
                    setActionError(null);
                    setActionOk(null);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, (0,_services_sharePointListService__WEBPACK_IMPORTED_MODULE_3__.deleteListItem)(client, listKey, row.id)];
                case 2:
                    _a.sent();
                    setActionOk("Deleted item #" + row.id);
                    onRefresh();
                    return [3 /*break*/, 5];
                case 3:
                    e_3 = _a.sent();
                    setActionError(e_3 instanceof Error ? e_3.message : "Delete failed.");
                    return [3 /*break*/, 5];
                case 4:
                    setBusy(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var exportRows = rows.map(function (r) { return r.cells; });
    var formKeys = Object.keys(form).sort();
    var previewSrc = localPreviewUrl || logoPreview;
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].panel },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].toolbar },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", onClick: openCreate, disabled: busy }, "New"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", onClick: function () { return (0,_services_exportService__WEBPACK_IMPORTED_MODULE_5__.exportTableAsCsv)(title, headers, exportRows); }, disabled: loading || rows.length === 0 }, "Export CSV"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", onClick: function () { return (0,_services_exportService__WEBPACK_IMPORTED_MODULE_5__.exportTableAsExcel)(title, headers, exportRows); }, disabled: loading || rows.length === 0 }, "Export Excel"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", onClick: onRefresh, disabled: busy || loading }, "Refresh")),
        actionError && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].error }, actionError),
        actionOk && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].success }, actionOk),
        loading && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].muted }, "Loading\u2026"),
        error && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].error }, error),
        !loading && !error && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].tableWrap },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("table", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].table },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("thead", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].actionsCol }, "Actions"),
                        headers.map(function (h) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", { key: h }, h)); }))),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("tbody", null, rows.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { colSpan: (headers.length || 0) + 1, className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].muted }, "No rows found."))) : (rows.map(function (row) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", { key: row.id },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].actionsCol },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].linkBtn, onClick: function () {
                                openEdit(row).catch(function () { return undefined; });
                            }, disabled: busy }, "Edit"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].linkBtnDanger, onClick: function () {
                                remove(row).catch(function () { return undefined; });
                            }, disabled: busy }, "Delete")),
                    row.cells.map(function (c, i) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { key: row.id + "-" + i }, c)); }))); })))))),
        drawerOpen && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].drawerBackdrop },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].drawer, role: "dialog", "aria-modal": "true" },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].drawerHeader },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].title },
                        editingId ? "Edit item #" + editingId : "New item",
                        " \u2014",
                        " ",
                        listLabel),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].linkBtn, onClick: function () {
                            clearLogoSelection();
                            setDrawerOpen(false);
                        }, disabled: busy }, "Close")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].muted }, "Fill Company Name (required). Blank fields are skipped. Lookup columns: enter the SharePoint item ID."),
                actionError && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].error }, actionError),
                supportsCompanyLogo && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].logoBox },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].fieldLabel }, "Company Logo"),
                    previewSrc ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("img", { src: previewSrc, alt: "Company logo preview", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].logoPreview })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].muted }, "No logo yet")),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("input", { type: "file", accept: "image/*", disabled: busy, onChange: function (e) { return onLogoChosen(e.target.files); } }),
                    logoFile && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].muted },
                        "Selected: ",
                        logoFile.name,
                        " \u2014 will upload on Save")),
                    logoFile && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].linkBtn, disabled: busy, onClick: clearLogoSelection }, "Clear selection")))),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].formGrid }, formKeys.map(function (key) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("label", { key: key, className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].field },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].fieldLabel },
                        key,
                        isBooleanField(key) ? " (yes/no)" : "",
                        original[key + "Id"] !== undefined
                            ? " (lookup ID)"
                            : ""),
                    isBooleanField(key) ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("select", { value: form[key] || "false", onChange: function (e) {
                            var _a;
                            return setForm(__assign(__assign({}, form), (_a = {}, _a[key] = e.target.value, _a)));
                        }, disabled: busy },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("option", { value: "true" }, "Yes"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("option", { value: "false" }, "No"))) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("input", { type: "text", value: form[key] || "", onChange: function (e) {
                            var _a;
                            return setForm(__assign(__assign({}, form), (_a = {}, _a[key] = e.target.value, _a)));
                        }, disabled: busy })))); })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].drawerFooter },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", onClick: save, disabled: busy }, busy
                        ? "Saving…"
                        : editingId
                            ? "Save changes"
                            : "Create"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", onClick: function () {
                            clearLogoSelection();
                            setDrawerOpen(false);
                        }, disabled: busy }, "Cancel")))))));
};


/***/ }),

/***/ 303:
/*!********************************************!*\
  !*** ./lib/shared/ui/AdminHubDashboard.js ***!
  \********************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AdminHubDashboard: () => (/* binding */ AdminHubDashboard)
/* harmony export */ });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 650);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./portal.module.scss */ 256);


var ACTION_TILES = [
    { id: "companies", title: "Companies", hint: "Company List", icon: "C" },
    { id: "workforce", title: "Workforce", hint: "Candidates", icon: "W" },
    { id: "documents", title: "Documents", hint: "Upload & visibility", icon: "D" },
    { id: "events", title: "Calendar", hint: "Bookings", icon: "B" },
    { id: "training-matrix", title: "Matrix", hint: "Training expiries", icon: "M" },
    { id: "training-records", title: "Registers", hint: "NPORS · EUSR · more", icon: "R" },
    { id: "offers", title: "Offers", hint: "Promotions", icon: "O" },
    { id: "permissions", title: "Permissions", hint: "Portal access", icon: "P" },
];
var RESOURCE_TILES = [
    {
        id: "training-matrix",
        title: "Training Matrix",
        description: "Wide expiry grid and sync",
        tone: "lime",
    },
    {
        id: "training-records",
        title: "Training Registers",
        description: "NPORS, EUSR, Streetworks, In-House",
        tone: "charcoal",
    },
    {
        id: "documents",
        title: "Customer Documents",
        description: "Folders, uploads, visibility",
        tone: "forest",
    },
    {
        id: "events",
        title: "Calendar / Bookings",
        description: "Events and Outlook sync",
        tone: "slate",
    },
    {
        id: "permissions",
        title: "Permissions",
        description: "Who can access the portals",
        tone: "moss",
    },
    {
        id: "logs",
        title: "Audit Log",
        description: "Portal activity history",
        tone: "ink",
    },
];
function displayNameFromEmail(email) {
    var local = (email.split("@")[0] || email).trim();
    var parts = local.split(/[.\-_+\s]+/).filter(Boolean);
    if (parts.length === 0)
        return "Admin";
    return parts
        .map(function (part) { return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase(); })
        .join(" ");
}
var TONE_CLASS = {
    lime: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubTone_lime,
    charcoal: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubTone_charcoal,
    forest: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubTone_forest,
    slate: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubTone_slate,
    moss: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubTone_moss,
    ink: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubTone_ink,
};
function friendlyStatLabel(key) {
    var map = {
        companies: "Active companies",
        workforce: "Active candidates",
        matrix: "Matrix rows",
        documents: "Documents",
        events: "Upcoming bookings",
        nvq: "NVQs",
        permissions: "Permissions",
    };
    return map[key] || key;
}
var AdminHubDashboard = function (props) {
    var userEmail = props.userEmail, counts = props.counts, onNavigate = props.onNavigate;
    var welcomeName = displayNameFromEmail(userEmail);
    var countKeys = Object.keys(counts);
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubPage },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHero, "aria-label": "Welcome" },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroInner },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroEyebrow }, "PAVE Training \u00B7 Admin"),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("h1", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroTitle },
                    "Welcome, ",
                    welcomeName),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroSubtitle }, "Operations hub for companies, workforce, matrix, bookings, and customer access \u2014 powered by SharePoint."),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroPanel },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroPanelLabel }, "Operations"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroPanelHint }, "Use the tiles below for everyday admin tasks. All lists open with full create, edit, delete, and export \u2014 same SharePoint data as the Next.js admin.")),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroPanelLabel }, "Signed in"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubHeroPanelHint }, userEmail))))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubActionGrid, "aria-label": "Quick actions" }, ACTION_TILES.map(function (tile) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: tile.id, type: "button", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubActionTile, onClick: function () { return onNavigate(tile.id); }, "aria-label": "Open ".concat(tile.title) },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubActionIcon, "aria-hidden": true }, tile.icon),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubActionTitle }, tile.title),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubActionHint }, tile.hint))); })),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubResources, "aria-label": "Top resources" },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubResourcesHeader },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubResourcesTitle }, "Top resources"),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubResourcesSubtitle }, "Jump into the lists you use most.")),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubResourceGrid }, RESOURCE_TILES.map(function (tile) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: tile.id, type: "button", className: "".concat(_portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].hubResourceTile, " ").concat(TONE_CLASS[tile.tone]), onClick: function () { return onNavigate(tile.id); }, "aria-label": "Open ".concat(tile.title) },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("strong", null, tile.title),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", null, tile.description))); }))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].stats, "aria-label": "Dashboard statistics" }, countKeys.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].muted }, "Loading SharePoint counts\u2026")) : (countKeys.map(function (key) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("article", { key: key, className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].stat },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].statValue }, counts[key]),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_1__["default"].statLabel }, friendlyStatLabel(key)))); })))));
};


/***/ }),

/***/ 787:
/*!*********************************************!*\
  !*** ./lib/shared/ui/CustomerPortalView.js ***!
  \*********************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   CustomerPortalView: () => (/* binding */ CustomerPortalView)
/* harmony export */ });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 650);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _services_permissionService__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../services/permissionService */ 942);
/* harmony import */ var _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./customerPortal.module.scss */ 276);



// eslint-disable-next-line @typescript-eslint/no-var-requires
var paveLogo = __webpack_require__(/*! ../assets/pave-logo.png */ 659);
var DESKTOP_NAV = [
    { id: "training-matrix", label: "Training Matrix" },
    { id: "dashboard", label: "Dashboard" },
    { id: "candidates", label: "Candidates" },
    { id: "documents", label: "Documents" },
    { id: "nvq-progress", label: "NVQ Progress" },
    { id: "events", label: "Events" },
    { id: "offers", label: "Offers" },
    { id: "support", label: "Support" },
];
var MORE_NAV = [
    { id: "training-records", label: "Training Records" },
    { id: "candidates", label: "Candidates" },
    { id: "offers", label: "Offers" },
    { id: "support", label: "Support" },
];
function initials(name) {
    var parts = name.trim().split(/\s+/).filter(Boolean);
    if (parts.length === 0)
        return "?";
    if (parts.length === 1)
        return parts[0].substring(0, 2).toUpperCase();
    return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
}
function statusKind(raw) {
    var s = (raw || "").toLowerCase();
    if (s.indexOf("missing") >= 0 ||
        s.indexOf("expired") >= 0 ||
        s.indexOf("review") >= 0 ||
        s.indexOf("fail") >= 0) {
        return "bad";
    }
    if (s.indexOf("expir") >= 0 ||
        s.indexOf("soon") >= 0 ||
        s.indexOf("due") >= 0 ||
        s.indexOf("attention") >= 0) {
        return "warn";
    }
    return "ok";
}
function statusLabel(raw) {
    var kind = statusKind(raw);
    if (kind === "bad")
        return raw.trim() || "Missing Data";
    if (kind === "warn")
        return raw.trim() || "Expiring Soon";
    if (!raw.trim() || raw.toLowerCase() === "compliant")
        return "Compliant";
    return raw.trim();
}
function IconPeople() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M16 11a4 4 0 1 0-4-4 4 4 0 0 0 4 4ZM8 13a4 4 0 1 0-4-4 4 4 0 0 0 4 4ZM16 13c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4ZM8 13c-.29 0-.62.02-.97.05C4.84 13.56 2 14.94 2 17v2h6", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round", strokeLinejoin: "round" })));
}
function IconMatrix() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M4 4h7v7H4V4Zm9 0h7v7h-7V4ZM4 13h7v7H4v-7Zm9 0h7v7h-7v-7Z", stroke: "currentColor", strokeWidth: "1.7" })));
}
function IconDocs() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M7 3h7l5 5v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M14 3v5h5", stroke: "currentColor", strokeWidth: "1.7" })));
}
function IconNvq() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M12 3 4 7v5c0 5 3.4 8.4 8 9 4.6-.6 8-4 8-9V7l-8-4Z", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "m9 12 2 2 4-4", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round" })));
}
function IconEvents() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("rect", { x: "3", y: "5", width: "18", height: "16", rx: "2", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M3 10h18M8 3v4M16 3v4", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round" })));
}
function IconOffers() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M12 3 4 7v4c0 4.5 3 8.2 8 10 5-1.8 8-5.5 8-10V7l-8-4Z", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M9 12h6", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round" })));
}
function IconSupport() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("circle", { cx: "12", cy: "12", r: "9", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M9.5 9.5a2.5 2.5 0 1 1 3.6 2.2c-.7.4-1.1.9-1.1 1.8M12 17h.01", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round" })));
}
function IconClock() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { width: "18", height: "18", viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("circle", { cx: "12", cy: "12", r: "8", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M12 8v5l3 2", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round" })));
}
function IconWarn() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { width: "18", height: "18", viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M12 4 3 19h18L12 4Z", stroke: "currentColor", strokeWidth: "1.7", strokeLinejoin: "round" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M12 10v4M12 16.5h.01", stroke: "currentColor", strokeWidth: "1.7", strokeLinecap: "round" })));
}
function IconFolder() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].docIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7Z", fill: "#F5C542", stroke: "#D4A017", strokeWidth: "1" })));
}
function IconPdf() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].docIcon, viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M7 3h7l5 5v13H7V3Z", fill: "#FEE2E2", stroke: "#DC2626", strokeWidth: "1.4" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M14 3v5h5", stroke: "#DC2626", strokeWidth: "1.4" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("text", { x: "8", y: "17", fontSize: "6", fontWeight: "700", fill: "#DC2626" }, "PDF")));
}
function IconPin() {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { width: "12", height: "12", viewBox: "0 0 24 24", fill: "none", "aria-hidden": "true" },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M12 21s7-5.2 7-11a7 7 0 1 0-14 0c0 5.8 7 11 7 11Z", stroke: "currentColor", strokeWidth: "1.7" }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("circle", { cx: "12", cy: "10", r: "2.2", stroke: "currentColor", strokeWidth: "1.7" })));
}
function navIcon(id) {
    switch (id) {
        case "dashboard":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconPeople, null);
        case "training-matrix":
        case "candidates":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconMatrix, null);
        case "documents":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconDocs, null);
        case "nvq-progress":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconNvq, null);
        case "events":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconEvents, null);
        case "offers":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconOffers, null);
        case "support":
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconSupport, null);
        default:
            return react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconMatrix, null);
    }
}
function parseEventDate(raw) {
    var d = new Date(raw);
    if (!isNaN(d.getTime())) {
        return {
            month: d.toLocaleString("en-GB", { month: "short" }).toUpperCase(),
            day: String(d.getDate()),
        };
    }
    return { month: "—", day: "—" };
}
/**
 * Presentational Customer Portal shell — matches PAVE design mock.
 * Receives already-loaded portal state; does not fetch.
 */
var CustomerPortalView = function (props) {
    var permission = props.permission, view = props.view, onNavigate = props.onNavigate, counts = props.counts, headers = props.headers, rows = props.rows, loading = props.loading, error = props.error, stub = props.stub, pageTitle = props.pageTitle, pageSubtitle = props.pageSubtitle, dashboard = props.dashboard;
    var _a = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), menuOpen = _a[0], setMenuOpen = _a[1];
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState(""), matrixQuery = _b[0], setMatrixQuery = _b[1];
    var companyName = permission.companyDisplayName || permission.userEmail || "Customer";
    var avatar = initials(companyName);
    var accessLabel = (0,_services_permissionService__WEBPACK_IMPORTED_MODULE_1__.accessScopeBadgeLabel)(permission);
    var downloadLabel = permission.canDownload
        ? "Downloads: Enabled"
        : "Downloads: Disabled";
    var totalCandidates = counts.workforce != null
        ? counts.workforce
        : dashboard && dashboard.counts.workforce != null
            ? dashboard.counts.workforce
            : 0;
    var upcomingEvents = counts.events != null
        ? counts.events
        : dashboard && dashboard.counts.events != null
            ? dashboard.counts.events
            : 0;
    var expiringSoon = counts.expiringSoon != null
        ? counts.expiringSoon
        : dashboard && dashboard.counts.expiringSoon != null
            ? dashboard.counts.expiringSoon
            : 0;
    var missingData = counts.missingData != null
        ? counts.missingData
        : dashboard && dashboard.counts.missingData != null
            ? dashboard.counts.missingData
            : 0;
    var sourceMatrix = dashboard && dashboard.matrixRows.length > 0
        ? dashboard.matrixRows
        : rows;
    var matrixPreview = react__WEBPACK_IMPORTED_MODULE_0__.useMemo(function () {
        var q = matrixQuery.trim().toLowerCase();
        var list = sourceMatrix;
        if (q) {
            list = list.filter(function (r) { return r.cells.join(" ").toLowerCase().indexOf(q) >= 0; });
        }
        return list.slice(0, 5);
    }, [sourceMatrix, matrixQuery]);
    var closeMenu = function () { return setMenuOpen(false); };
    var go = function (id) {
        onNavigate(id);
        closeMenu();
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].root },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("header", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].header },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].headerLeft },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("img", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].logo, src: paveLogo, alt: "PAVE Training" }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].headerCopy },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h1", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].portalTitle }, "PAVE Training Customer Portal"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].welcome },
                        "Welcome back,",
                        " ",
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].welcomeName }, permission.userEmail)),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tagline },
                        "Company: ",
                        companyName),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].accessBadges, "aria-label": "Access summary" },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].badge },
                            "Role: ",
                            permission.roleLabel),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].badge },
                            "Access: ",
                            accessLabel),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: permission.canDownload
                                ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].badgeOk
                                : _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].badgeMuted }, downloadLabel))),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].menuToggle, "aria-label": "Open menu", onClick: function () { return setMenuOpen(!menuOpen); } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { width: "18", height: "18", viewBox: "0 0 24 24", fill: "none" },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("path", { d: "M4 7h16M4 12h16M4 17h16", stroke: "currentColor", strokeWidth: "2", strokeLinecap: "round" })))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].userMenu, onClick: function () { return undefined; } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].userName }, permission.userEmail),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].chevron, "aria-hidden": "true" }, "\u25BE"),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].avatar }, avatar))),
        menuOpen && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].mobileNavDrawer }, DESKTOP_NAV.concat(MORE_NAV.filter(function (m) { return m.id === "candidates" || m.id === "training-records"; })).map(function (item) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: "m-" + item.id, type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navItem, " ").concat(view === item.id ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navItemActive : ""), onClick: function () { return go(item.id); } },
            navIcon(item.id),
            item.label)); }))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("nav", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navBar, "aria-label": "Customer portal" }, DESKTOP_NAV.map(function (item) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: item.id, type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navItem, " ").concat(view === item.id ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navItemActive : ""), onClick: function () { return go(item.id); } },
            navIcon(item.id),
            item.label)); })),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].body },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].mobileWelcome },
                "Welcome back,",
                " ",
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].welcomeName }, companyName)),
            view === "dashboard" ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(DashboardBody, { totalCandidates: totalCandidates, expiringSoon: expiringSoon, missingData: missingData, upcomingEvents: upcomingEvents, matrixRows: matrixPreview, matrixTotal: sourceMatrix.length || totalCandidates, matrixQuery: matrixQuery, onMatrixQuery: setMatrixQuery, onNavigate: go, loading: loading, error: error, documentTiles: dashboard && dashboard.documentTiles.length > 0
                    ? dashboard.documentTiles
                    : [], nvqRows: dashboard && dashboard.nvqRows.length > 0
                    ? dashboard.nvqRows
                    : [], eventRows: dashboard && dashboard.eventRows.length > 0
                    ? dashboard.eventRows
                    : [], offerCards: dashboard && dashboard.offerCards.length > 0
                    ? dashboard.offerCards
                    : [] })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(ListBody, { title: pageTitle, subtitle: pageSubtitle, stub: stub, headers: headers, rows: rows, loading: loading, error: error, view: view, onNavigate: go }))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("nav", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].bottomNav, "aria-label": "Mobile navigation" },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItem, " ").concat(view === "dashboard" ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItemActive : ""), onClick: function () { return go("dashboard"); } },
                navIcon("dashboard"),
                "Dashboard"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItem, " ").concat(view === "training-matrix" ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItemActive : ""), onClick: function () { return go("training-matrix"); } },
                navIcon("training-matrix"),
                "Matrix"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItem, " ").concat(view === "documents" ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItemActive : ""), onClick: function () { return go("documents"); } },
                navIcon("documents"),
                "Documents"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItem, " ").concat(view === "events" ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItemActive : ""), onClick: function () { return go("events"); } },
                navIcon("events"),
                "Events"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItem, " ").concat(menuOpen ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tabItemActive : ""), onClick: function () { return setMenuOpen(!menuOpen); } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("svg", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].navIcon, viewBox: "0 0 24 24", fill: "currentColor" },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("circle", { cx: "5", cy: "12", r: "2" }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("circle", { cx: "12", cy: "12", r: "2" }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("circle", { cx: "19", cy: "12", r: "2" })),
                "More"))));
};
var DashboardBody = function (p) {
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].stats },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statCard },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statLabel }, "Total Candidates"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statValue }, p.totalCandidates),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHint, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHintGood) }, "Company workforce")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIcon, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIconGreen) },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconPeople, null))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statCard },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statLabel }, "Expiring Soon"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statValue }, p.expiringSoon),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHint, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHintWarn) }, "Within 60 days")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIcon, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIconOrange) },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconClock, null))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statCard },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statLabel }, "Missing Data"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statValue }, p.missingData),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHint, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHintWarn) }, "Requires attention")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIcon, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIconRed) },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconWarn, null))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statCard },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statLabel }, "Upcoming Events"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statValue }, p.upcomingEvents),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHint, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statHintGood) }, "Next 30 days")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIcon, " ").concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statIconGreen) },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconEvents, null)))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].dashMain },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].card },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardHeader },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardTitle }, "Training Matrix"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardActions },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("input", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].searchInput, type: "search", placeholder: "Search candidates...", value: p.matrixQuery, onChange: function (e) { return p.onMatrixQuery(e.target.value); } }),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].ghostBtn, onClick: function () { return undefined; } }, "Filters"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].primaryBtn, onClick: function () { return p.onNavigate("training-matrix"); } }, "View full matrix"))),
                p.loading && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "Loading\u2026"),
                p.error && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].error }, p.error),
                !p.loading && !p.error && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(react__WEBPACK_IMPORTED_MODULE_0__.Fragment, null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(MatrixTable, { rows: p.matrixRows, condensed: false, emptyHint: "No training matrix rows for your company yet." }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tableFooter },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted },
                            "Showing 1 to ",
                            Math.min(5, p.matrixRows.length),
                            " of",
                            " ",
                            p.matrixTotal,
                            " candidates"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, onClick: function () { return p.onNavigate("training-matrix"); } }, "View full training matrix \u2192"))))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].dashSide },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].card },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardHeader },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardTitle }, "Customer Documents"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, onClick: function () { return p.onNavigate("documents"); } }, "View all")),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].docGrid }, p.documentTiles.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "No visible documents for your company.")) : (p.documentTiles.map(function (doc) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: doc.id, type: "button", className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].docTile, " ").concat(doc.kind === "pdf" ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pdfTile : ""), onClick: function () { return p.onNavigate("documents"); } },
                        doc.kind === "pdf" ? react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconPdf, null) : react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconFolder, null),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].docLabel }, doc.label),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].docMeta }, doc.meta))); })))),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].card },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardHeader },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardTitle }, "NVQ Progress"),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, onClick: function () { return p.onNavigate("nvq-progress"); } }, "View all")),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].nvqList }, p.nvqRows.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "No visible NVQ records.")) : (p.nvqRows.map(function (row) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { key: row.id, className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].nvqRow },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].nvqTop },
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].nvqName }, row.name),
                                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].nvqCourse }, row.course)),
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].nvqPct },
                                row.pct,
                                "%")),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].progressTrack },
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].progressFill, style: { width: row.pct + "%" } })))); })))))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].dashBottom },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].card },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardHeader },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardTitle }, "Upcoming Events"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, onClick: function () { return p.onNavigate("events"); } }, "View all")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].eventList }, p.eventRows.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "No upcoming visible events.")) : (p.eventRows.map(function (ev) {
                    var badge = parseEventDate(ev.dateRaw);
                    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: ev.id, type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].eventCard, onClick: function () { return p.onNavigate("events"); } },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].dateBadge },
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].dateMonth }, badge.month),
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].dateDay }, badge.day)),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", null,
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].eventTitle }, ev.title),
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].eventMeta }, ev.when),
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].eventMeta },
                                react__WEBPACK_IMPORTED_MODULE_0__.createElement(IconPin, null),
                                " ",
                                ev.where))));
                })))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].card },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardHeader },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].cardTitle }, "Offers & Promotions"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, onClick: function () { return p.onNavigate("offers"); } }, "View all")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].offerGrid }, p.offerCards.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "No visible offers right now.")) : (p.offerCards.map(function (offer) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: offer.id, type: "button", className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].offerCard, onClick: function () { return p.onNavigate("offers"); } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].offerBadge }, offer.badge),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].offerTitle }, offer.title),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: offer.code.indexOf("→") >= 0
                            ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].offerLink
                            : _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].offerCode }, offer.code))); })))))));
};
function MatrixTable(props) {
    var rows = props.rows, condensed = props.condensed, emptyHint = props.emptyHint;
    if (rows.length === 0) {
        return react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, emptyHint);
    }
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tableWrap },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("table", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].table },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("thead", null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Candidate"),
                    !condensed && react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Role"),
                    !condensed && react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Key Training"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Status"),
                    !condensed && react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Expiry Date"))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("tbody", null, rows.map(function (row) {
                var name = row.cells[0] || "—";
                var status = row.cells[1] || "";
                var expiry = row.cells[2] || "—";
                var kind = statusKind(status);
                var pillClass = kind === "bad"
                    ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pillBad
                    : kind === "warn"
                        ? _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pillWarn
                        : _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pillOk;
                return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", { key: row.id },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].candidateCell },
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].avatarSm }, initials(name)),
                            name)),
                    !condensed && react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "\u2014"),
                    !condensed && react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "\u2014"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, condensed ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pill, " ").concat(pillClass), title: statusLabel(status) }, kind === "ok" ? "✓" : kind === "warn" ? "⏱" : "!")) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: "".concat(_customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pill, " ").concat(pillClass) }, statusLabel(status)))),
                    !condensed && react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, expiry)));
            })))));
}
function ListBody(props) {
    var title = props.title, subtitle = props.subtitle, stub = props.stub, headers = props.headers, rows = props.rows, loading = props.loading, error = props.error, view = props.view;
    if (stub) {
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].listPage },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pageTitle }, title),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pageSubtitle }, subtitle),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].stubBox }, stub)));
    }
    var isMatrix = view === "training-matrix";
    var isDocuments = view === "documents";
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].listPage },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pageTitle }, title),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].pageSubtitle },
            subtitle,
            !loading ? " · " + rows.length + " rows" : ""),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].card },
            loading && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "Loading\u2026"),
            error && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].error }, error),
            !loading && !error && isMatrix && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(MatrixTable, { rows: rows.slice(0, 50), condensed: false, emptyHint: "No matrix rows found for your company." })),
            !loading && !error && isDocuments && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(DocumentsTable, { rows: rows })),
            !loading && !error && !isMatrix && !isDocuments && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tableWrap },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("table", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].table },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("thead", null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", null, headers.map(function (h) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", { key: h }, h)); }))),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("tbody", null, rows.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { colSpan: headers.length || 1, className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "No rows found."))) : (rows.map(function (row) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", { key: row.id }, row.cells.map(function (c, i) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", { key: row.id + "-" + i }, c)); }))); })))))))));
}
function DocumentsTable(props) {
    var rows = props.rows;
    if (rows.length === 0) {
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "No documents have been shared with your account yet."));
    }
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].tableWrap },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("table", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].table },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("thead", null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Document Name"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Document Type"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Candidate Name"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Modified Date"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "View"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("th", null, "Download"))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("tbody", null, rows.map(function (row) {
                var name = row.cells[0] || "—";
                var type = row.cells[1] || "—";
                var candidate = row.cells[2] || "—";
                var modified = row.cells[3] || "—";
                var viewUrl = (row.fields && row.fields.__docViewUrl) ||
                    row.cells[4] ||
                    "";
                var downloadUrl = (row.fields && row.fields.__docDownloadUrl) ||
                    row.cells[5] ||
                    "";
                var canDownload = Boolean(row.fields && row.fields.__docCanDownload);
                return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("tr", { key: row.id },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, name),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, type),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, candidate),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, modified),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, viewUrl ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("a", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, href: viewUrl, target: "_blank", rel: "noreferrer" }, "View")) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "\u2014"))),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("td", null, canDownload && downloadUrl ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("a", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].linkBtn, href: downloadUrl, target: "_blank", rel: "noreferrer", download: true }, "Download")) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _customerPortal_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].muted }, "\u2014")))));
            })))));
}


/***/ }),

/***/ 698:
/*!**************************************!*\
  !*** ./lib/shared/ui/PortalShell.js ***!
  \**************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   PortalShell: () => (/* binding */ PortalShell)
/* harmony export */ });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 650);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../schema/sharepointSchema */ 784);
/* harmony import */ var _services_customerAccessService__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../services/customerAccessService */ 410);
/* harmony import */ var _services_permissionService__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../services/permissionService */ 942);
/* harmony import */ var _services_portalDataService__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../services/portalDataService */ 350);
/* harmony import */ var _services_sharePointListService__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../services/sharePointListService */ 161);
/* harmony import */ var _AdminDataTable__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./AdminDataTable */ 179);
/* harmony import */ var _AdminHubDashboard__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./AdminHubDashboard */ 303);
/* harmony import */ var _CustomerPortalView__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./CustomerPortalView */ 787);
/* harmony import */ var _nav__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./nav */ 700);
/* harmony import */ var _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ./portal.module.scss */ 256);
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};











var ADMIN_VIEWS = {
    dashboard: {
        title: "Dashboard",
        subtitle: "Operations overview from SharePoint lists",
    },
    companies: {
        title: "Companies",
        subtitle: "Company List",
        listKey: "company",
        headers: ["Company", "Number", "Email", "Status"],
        columns: ["CompanyName", "CompanyNumber", "Email", "Status"],
    },
    workforce: {
        title: "Workforce",
        subtitle: "Workforce List",
        listKey: "workforce",
        headers: ["Candidate", "Company", "Department", "Status"],
        columns: ["CandidateName", "CompanyName", "Department", "Status"],
    },
    "training-matrix": {
        title: "Training Matrix",
        subtitle: "Matrix rows and expiry dates",
        listKey: "trainingMatrix",
        headers: ["Candidate", "Company", "Status", "Next expiry"],
        columns: [
            "CandidateName",
            "Company_x0020_Name",
            "OverallStatus",
            "NextExpiryDate",
        ],
    },
    "training-records": {
        title: "Training Records",
        subtitle: "Choose a register",
    },
    nvq: {
        title: "NVQ",
        subtitle: "NVQ Register",
        listKey: "nvqRegister",
        headers: ["Candidate", "Title", "Stage", "Company"],
        columns: [
            "CandidateName",
            "NvqTitle",
            "StageofNvq",
            "Company_x0020_Name",
        ],
    },
    documents: {
        title: "Documents",
        subtitle: "Customer Documents library",
        listKey: "customerDocuments",
        headers: ["Name", "Company", "Type", "Visible"],
        columns: ["FileLeafRef", "Company", "DocumentType", "CustomerVisible"],
    },
    events: {
        title: "Events",
        subtitle: "Events calendar list",
        listKey: "events",
        headers: ["Title", "Company", "Start", "Visible"],
        columns: ["Title", "EventCompany", "EventDate", "Customer_x0020_Visible"],
    },
    offers: {
        title: "Offers",
        subtitle: "Offers / Promotions",
        listKey: "offersPromotions",
        headers: ["Title", "Category", "Status", "Visible"],
        columns: ["Title", "Category", "Status", "CustomerVisible"],
    },
    permissions: {
        title: "Permissions",
        subtitle: "Portal access control",
        listKey: "permissions",
        headers: ["Email", "Role", "Status", "Scope"],
        columns: ["UserEmail", "RoleType", "Status", "AccessScope"],
    },
    automation: {
        title: "Automation",
        subtitle: "Automation rules and sync controls",
        stub: "Automation rules and Outlook sync controls will connect here (same concept as Next.js /admin/automation).",
    },
    logs: {
        title: "Logs",
        subtitle: "Training Manager Logs",
        listKey: "trainingManagerLogs",
        headers: ["Title", "User", "List", "Timestamp"],
        columns: ["Title", "User_x0020_Email", "ListName", "Timestamp"],
    },
};
var CUSTOMER_VIEWS = {
    dashboard: {
        title: "Dashboard",
        subtitle: "Your company training overview",
    },
    "training-matrix": {
        title: "Training Matrix",
        subtitle: "Your company's matrix rows",
        listKey: "trainingMatrix",
        headers: ["Candidate", "Status", "Next expiry"],
        columns: ["CandidateName", "OverallStatus", "NextExpiryDate"],
        companyField: "Company_x0020_Name",
        customerScoped: true,
    },
    candidates: {
        title: "Candidates",
        subtitle: "Workforce for your company",
        listKey: "workforce",
        headers: ["Candidate", "Department", "Status"],
        columns: ["CandidateName", "Department", "Status"],
        companyField: "CompanyName",
        customerScoped: true,
    },
    "training-records": {
        title: "Training Records",
        subtitle: "Choose a register",
    },
    "nvq-progress": {
        title: "NVQ Progress",
        subtitle: "Visible NVQ records for your company",
        listKey: "nvqRegister",
        headers: ["Candidate", "Title", "Stage"],
        columns: ["CandidateName", "NvqTitle", "StageofNvq"],
        companyField: "Company_x0020_Name",
        visibleField: "CustomerVisible",
        customerScoped: true,
    },
    documents: {
        title: "Documents",
        subtitle: "Certificates, card scans, NVQs, brochures and shared files",
        listKey: "customerDocuments",
        headers: [
            "Document Name",
            "Document Type",
            "Candidate Name",
            "Modified Date",
            "View",
            "Download",
        ],
        columns: [
            "FileLeafRef",
            "DocumentType",
            "Candidate",
            "Modified",
            "__view",
            "__download",
        ],
        companyField: "Company",
        visibleField: "CustomerVisible",
        customerScoped: true,
    },
    events: {
        title: "Events",
        subtitle: "Visible upcoming events",
        listKey: "events",
        headers: ["Title", "Start", "Address"],
        columns: ["Title", "EventDate", "TrainingAddress"],
        companyField: "EventCompany",
        visibleField: "Customer_x0020_Visible",
        customerScoped: true,
    },
    offers: {
        title: "Offers",
        subtitle: "Visible offers and promotions",
        listKey: "offersPromotions",
        headers: ["Title", "Category", "Status"],
        columns: ["Title", "Category", "Status"],
        visibleField: "CustomerVisible",
        customerScoped: true,
    },
    support: {
        title: "Support",
        subtitle: "Contact your PAVE training manager",
        stub: "For renewals, certificates, or portal access, contact your assigned Training Manager. Include your company name and login email.",
    },
};
var RECORD_SUBNAV = [
    { id: "npors", label: "NPORS", listKey: "nporsRegister" },
    { id: "eusr", label: "EUSR", listKey: "eusrRegister" },
    {
        id: "streetworks",
        label: "Streetworks",
        listKey: "nrswaRegister",
    },
    {
        id: "in-house",
        label: "In-House",
        listKey: "inHouseCertificates",
    },
];
var PortalShell = function (props) {
    var mode = props.mode, spHttpClient = props.spHttpClient, webUrl = props.webUrl, userEmail = props.userEmail, isSiteAdmin = props.isSiteAdmin;
    var client = react__WEBPACK_IMPORTED_MODULE_0__.useMemo(function () { return ({ spHttpClient: spHttpClient, webUrl: webUrl }); }, [spHttpClient, webUrl]);
    var resolvedEmail = (0,_services_sharePointListService__WEBPACK_IMPORTED_MODULE_5__.normalizeSharePointUserEmail)(userEmail);
    var _a = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), permission = _a[0], setPermission = _a[1];
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), permError = _b[0], setPermError = _b[1];
    var _c = react__WEBPACK_IMPORTED_MODULE_0__.useState(true), loadingPerm = _c[0], setLoadingPerm = _c[1];
    var _d = react__WEBPACK_IMPORTED_MODULE_0__.useState(mode === "customer" ? "training-matrix" : "dashboard"), view = _d[0], setView = _d[1];
    var _e = react__WEBPACK_IMPORTED_MODULE_0__.useState("npors"), recordTab = _e[0], setRecordTab = _e[1];
    var _f = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), rows = _f[0], setRows = _f[1];
    var _g = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), headers = _g[0], setHeaders = _g[1];
    var _h = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), columns = _h[0], setColumns = _h[1];
    var _j = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), activeListKey = _j[0], setActiveListKey = _j[1];
    var _k = react__WEBPACK_IMPORTED_MODULE_0__.useState({}), counts = _k[0], setCounts = _k[1];
    var _l = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), customerDash = _l[0], setCustomerDash = _l[1];
    var _m = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), loading = _m[0], setLoading = _m[1];
    var _o = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), error = _o[0], setError = _o[1];
    var _p = react__WEBPACK_IMPORTED_MODULE_0__.useState(0), refreshTick = _p[0], setRefreshTick = _p[1];
    var _q = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), mobileNavOpen = _q[0], setMobileNavOpen = _q[1];
    var nav = mode === "admin" ? _nav__WEBPACK_IMPORTED_MODULE_9__.ADMIN_NAV : _nav__WEBPACK_IMPORTED_MODULE_9__.CUSTOMER_NAV;
    var views = mode === "admin" ? ADMIN_VIEWS : CUSTOMER_VIEWS;
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        var cancelled = false;
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var profile, e_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        setLoadingPerm(true);
                        setPermError(null);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        return [4 /*yield*/, (0,_services_permissionService__WEBPACK_IMPORTED_MODULE_3__.getActivePermissionByEmail)(client, resolvedEmail || userEmail)];
                    case 2:
                        profile = _a.sent();
                        // Site collection admins can always use the Admin web part.
                        if (mode === "admin" &&
                            isSiteAdmin &&
                            (!profile || !profile.canAccessAdmin)) {
                            profile = (0,_services_permissionService__WEBPACK_IMPORTED_MODULE_3__.siteAdminPermissionProfile)(resolvedEmail || userEmail);
                        }
                        if (!cancelled) {
                            setPermission(profile);
                        }
                        return [3 /*break*/, 5];
                    case 3:
                        e_1 = _a.sent();
                        if (!cancelled) {
                            setPermError(e_1 instanceof Error ? e_1.message : "Failed to load permissions.");
                        }
                        return [3 /*break*/, 5];
                    case 4:
                        if (!cancelled) {
                            setLoadingPerm(false);
                        }
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        }); })();
        return function () {
            cancelled = true;
        };
    }, [client, userEmail, resolvedEmail, mode, isSiteAdmin]);
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        if (!permission) {
            return;
        }
        if (mode === "admin" && !permission.canAccessAdmin) {
            return;
        }
        if (mode === "customer" && !permission.canAccessCustomer) {
            return;
        }
        var cancelled = false;
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var companyName, dash, scopedMatrix, c, tab, cols_1, hdrs, companyField, visibleField, schema, fields, fields, fields, fields, fields, data_1, scoped_1, spec_1, docsRows, cols, headerLabels, schema, useLookupId, data, scoped, e_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        setLoading(true);
                        setError(null);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 12, 13, 14]);
                        companyName = mode === "customer" ? permission.companyDisplayName : undefined;
                        if (!(view === "dashboard")) return [3 /*break*/, 6];
                        if (!(mode === "customer" && permission.companyDisplayName)) return [3 /*break*/, 3];
                        return [4 /*yield*/, (0,_services_portalDataService__WEBPACK_IMPORTED_MODULE_4__.loadCustomerDashboardData)(client, permission.companyDisplayName, permission.companyId)];
                    case 2:
                        dash = _a.sent();
                        if (!cancelled) {
                            scopedMatrix = (0,_services_customerAccessService__WEBPACK_IMPORTED_MODULE_2__.filterPortalRowsByAccess)(dash.matrixRows, permission);
                            setCustomerDash(__assign(__assign({}, dash), { matrixRows: scopedMatrix, counts: __assign(__assign({}, dash.counts), { workforce: scopedMatrix.length, matrix: scopedMatrix.length }) }));
                            setCounts(__assign(__assign({}, dash.counts), { workforce: scopedMatrix.length, matrix: scopedMatrix.length }));
                            setRows(scopedMatrix);
                            setHeaders(["Candidate", "Status", "Next expiry"]);
                            setColumns([
                                "CandidateName",
                                "OverallStatus",
                                "NextExpiryDate",
                            ]);
                            setActiveListKey(null);
                        }
                        return [3 /*break*/, 5];
                    case 3: return [4 /*yield*/, (0,_services_portalDataService__WEBPACK_IMPORTED_MODULE_4__.loadDashboardCounts)(client, companyName)];
                    case 4:
                        c = _a.sent();
                        if (!cancelled) {
                            setCustomerDash(null);
                            setCounts(c);
                            setRows([]);
                            setHeaders([]);
                            setColumns([]);
                            setActiveListKey(null);
                        }
                        _a.label = 5;
                    case 5: return [2 /*return*/];
                    case 6:
                        if (!(view === "training-records")) return [3 /*break*/, 8];
                        tab = RECORD_SUBNAV.filter(function (t) { return t.id === recordTab; })[0];
                        hdrs = void 0;
                        companyField = "CompanyName";
                        visibleField = "CustomerVisible";
                        if (mode === "admin") {
                            schema = (0,_services_portalDataService__WEBPACK_IMPORTED_MODULE_4__.getAdminSchemaColumns)(tab.listKey);
                            cols_1 = schema.columns;
                            hdrs = schema.headers;
                            fields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.getSharePointFields)(tab.listKey);
                            companyField = fields.companyName;
                            visibleField = fields.customerVisible;
                        }
                        else if (tab.listKey === "nporsRegister") {
                            fields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.getSharePointFields)("nporsRegister");
                            cols_1 = [
                                fields.candidateName,
                                fields.nporsCategory,
                                fields.expiry,
                                fields.trainingOutcome,
                            ];
                            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
                            companyField = fields.companyName;
                            visibleField = fields.customerVisible;
                        }
                        else if (tab.listKey === "eusrRegister") {
                            fields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.getSharePointFields)("eusrRegister");
                            cols_1 = [
                                fields.candidateName,
                                fields.eusrCategory,
                                fields.expiry,
                                fields.trainingOutcome,
                            ];
                            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
                            companyField = fields.companyName;
                            visibleField = fields.customerVisible;
                        }
                        else if (tab.listKey === "nrswaRegister") {
                            fields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.getSharePointFields)("nrswaRegister");
                            cols_1 = [
                                fields.candidateName,
                                fields.course,
                                fields.expiryDate,
                                fields.trainingOutcome,
                            ];
                            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
                            companyField = fields.companyName;
                            visibleField = fields.customerVisible;
                        }
                        else {
                            fields = (0,_schema_sharepointSchema__WEBPACK_IMPORTED_MODULE_1__.getSharePointFields)("inHouseCertificates");
                            cols_1 = [
                                fields.candidateName,
                                fields.courseCategory,
                                fields.expiryDate,
                                fields.trainingOutcome,
                            ];
                            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
                            companyField = fields.companyName;
                            visibleField = fields.customerVisible;
                        }
                        return [4 /*yield*/, (0,_services_portalDataService__WEBPACK_IMPORTED_MODULE_4__.loadPortalListRows)(client, tab.listKey, {
                                columns: cols_1,
                                companyName: mode === "customer" ? permission.companyDisplayName : undefined,
                                companyFieldInternalName: mode === "customer" ? companyField : undefined,
                                customerVisibleOnly: mode === "customer",
                                visibleFieldInternalName: mode === "customer" ? visibleField : undefined,
                                loadAll: mode === "admin",
                            })];
                    case 7:
                        data_1 = _a.sent();
                        scoped_1 = mode === "customer"
                            ? (0,_services_customerAccessService__WEBPACK_IMPORTED_MODULE_2__.filterPortalRowsByAccess)(data_1, permission)
                            : data_1;
                        if (!cancelled) {
                            setHeaders(hdrs);
                            setColumns(cols_1);
                            setActiveListKey(tab.listKey);
                            setRows(scoped_1);
                        }
                        return [2 /*return*/];
                    case 8:
                        spec_1 = views[view];
                        if (!spec_1 || spec_1.stub || !spec_1.listKey) {
                            if (!cancelled) {
                                setRows([]);
                                setHeaders([]);
                                setColumns([]);
                                setActiveListKey(null);
                            }
                            return [2 /*return*/];
                        }
                        if (!(mode === "customer" && spec_1.listKey === "customerDocuments")) return [3 /*break*/, 10];
                        return [4 /*yield*/, (0,_services_customerAccessService__WEBPACK_IMPORTED_MODULE_2__.loadCustomerDocuments)(client, permission)];
                    case 9:
                        docsRows = _a.sent();
                        if (!cancelled) {
                            setHeaders(spec_1.headers || []);
                            setColumns(spec_1.columns || []);
                            setActiveListKey("customerDocuments");
                            setRows((0,_services_customerAccessService__WEBPACK_IMPORTED_MODULE_2__.documentRowsToPortalTable)(docsRows));
                        }
                        return [2 /*return*/];
                    case 10:
                        cols = spec_1.columns;
                        headerLabels = spec_1.headers;
                        if (mode === "admin" && spec_1.listKey) {
                            schema = (0,_services_portalDataService__WEBPACK_IMPORTED_MODULE_4__.getAdminSchemaColumns)(spec_1.listKey);
                            cols = schema.columns;
                            headerLabels = schema.headers;
                        }
                        if (!cols || !headerLabels) {
                            if (!cancelled) {
                                setRows([]);
                                setHeaders([]);
                                setColumns([]);
                                setActiveListKey(null);
                            }
                            return [2 /*return*/];
                        }
                        useLookupId = mode === "customer" &&
                            (spec_1.listKey === "events" || spec_1.listKey === "customerDocuments") &&
                            permission.companyId &&
                            permission.companyId !== "0";
                        return [4 /*yield*/, (0,_services_portalDataService__WEBPACK_IMPORTED_MODULE_4__.loadPortalListRows)(client, spec_1.listKey, {
                                columns: cols,
                                companyName: mode === "customer" && spec_1.customerScoped && !useLookupId
                                    ? permission.companyDisplayName
                                    : undefined,
                                companyFieldInternalName: mode === "customer" && !useLookupId ? spec_1.companyField : undefined,
                                companyId: useLookupId ? permission.companyId : undefined,
                                companyIdFieldInternalName: useLookupId
                                    ? spec_1.listKey === "events"
                                        ? "EventCompanyId"
                                        : "CompanyId"
                                    : undefined,
                                customerVisibleOnly: Boolean(mode === "customer" && spec_1.visibleField),
                                visibleFieldInternalName: spec_1.visibleField,
                                loadAll: mode === "admin",
                            })];
                    case 11:
                        data = _a.sent();
                        scoped = mode === "customer" &&
                            (spec_1.listKey === "workforce" ||
                                spec_1.listKey === "trainingMatrix" ||
                                spec_1.listKey === "nvqRegister")
                            ? (0,_services_customerAccessService__WEBPACK_IMPORTED_MODULE_2__.filterPortalRowsByAccess)(data, permission)
                            : data;
                        if (!cancelled) {
                            setHeaders(headerLabels);
                            setColumns(cols);
                            setActiveListKey(spec_1.listKey);
                            setRows(scoped);
                        }
                        return [3 /*break*/, 14];
                    case 12:
                        e_2 = _a.sent();
                        if (!cancelled) {
                            setError(e_2 instanceof Error ? e_2.message : "Failed to load data.");
                            setRows([]);
                        }
                        return [3 /*break*/, 14];
                    case 13:
                        if (!cancelled) {
                            setLoading(false);
                        }
                        return [7 /*endfinally*/];
                    case 14: return [2 /*return*/];
                }
            });
        }); })();
        return function () {
            cancelled = true;
        };
    }, [permission, view, recordTab, mode, client, views, refreshTick]);
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        setMobileNavOpen(false);
    }, [view]);
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        if (!mobileNavOpen) {
            return;
        }
        var previous = document.body.style.overflow;
        document.body.style.overflow = "hidden";
        return function () {
            document.body.style.overflow = previous;
        };
    }, [mobileNavOpen]);
    if (loadingPerm) {
        return react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted }, "Checking permissions\u2026");
    }
    if (permError) {
        return react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].error }, permError);
    }
    if (!permission) {
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].panel },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].title }, "Access denied"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted },
                "No active Permissions List row for ",
                resolvedEmail || userEmail,
                ". Add an Active Admin / Training Manager / Supervisor / Candidate entry in SharePoint Permissions List with matching User Email.")));
    }
    if (mode === "admin" && !permission.canAccessAdmin) {
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].panel },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].title }, "Admin access required"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted },
                "Signed in as ",
                permission.userEmail,
                " (",
                permission.roleLabel,
                "). This account can use the Customer Portal web part, but not Admin."),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted }, "PAVE internal Admin needs Role Type = Admin (or Training Manager for legacy access). Site collection admins can also open Admin after refresh.")));
    }
    if (mode === "customer" && !permission.canAccessCustomer) {
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].panel },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].title }, "Customer access required"),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted },
                "Signed in as ",
                permission.userEmail,
                " with role ",
                permission.roleLabel,
                ". Use the Admin Portal web part for PAVE internal admin accounts.")));
    }
    var spec = views[view];
    // Customer portal — design presentation layer
    if (mode === "customer") {
        if (!permission.companyDisplayName) {
            return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].panel },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].title }, "Company not resolved"),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted },
                    "Your Permissions List row is Active, but the Company lookup could not be resolved to a company name (CompanyId",
                    " ",
                    permission.companyId,
                    "). Check the Company field on your Permissions List item.")));
        }
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_CustomerPortalView__WEBPACK_IMPORTED_MODULE_8__.CustomerPortalView, { permission: permission, view: view, onNavigate: function (id) { return setView(id); }, counts: counts, headers: headers, rows: rows, loading: loading, error: error, stub: spec.stub, pageTitle: spec.title, pageSubtitle: spec.subtitle, dashboard: customerDash }));
    }
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: "".concat(_portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].shell, " ").concat(_portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminShell) },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("header", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminTopNav },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminTopNavBar },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminBrandBlock },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].brand }, "PAVE HUB"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].tagline }, "Admin operations")),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("nav", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminNavDesktop, "aria-label": "Admin" }, nav.map(function (item) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: item.id, type: "button", className: "".concat(_portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminNavLink, " ").concat(view === item.id ? _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminNavLinkActive : ""), onClick: function () { return setView(item.id); } }, item.label)); })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminTopNavTrailing },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminUserChip, title: permission.userEmail }, permission.userEmail),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { type: "button", className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMenuToggle, "aria-expanded": mobileNavOpen, "aria-controls": "pave-admin-mobile-nav", onClick: function () { return setMobileNavOpen(function (open) { return !open; }); } }, mobileNavOpen ? "Close" : "Menu"))),
            mobileNavOpen ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileScrim, onClick: function () { return setMobileNavOpen(false); }, "aria-hidden": true })) : null,
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { id: "pave-admin-mobile-nav", className: "".concat(_portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileDrawer, " ").concat(mobileNavOpen ? _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileDrawerOpen : ""), hidden: !mobileNavOpen },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileTitle }, "All options"),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("nav", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileNav, "aria-label": "Admin mobile" }, nav.map(function (item) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: item.id, type: "button", className: "".concat(_portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileLink, " ").concat(view === item.id ? _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].adminMobileLinkActive : ""), onClick: function () {
                        setView(item.id);
                        setMobileNavOpen(false);
                    } }, item.label)); })))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("main", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].main }, view === "dashboard" ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_AdminHubDashboard__WEBPACK_IMPORTED_MODULE_7__.AdminHubDashboard, { userEmail: permission.userEmail, counts: counts, onNavigate: function (id) { return setView(id); } })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(react__WEBPACK_IMPORTED_MODULE_0__.Fragment, null,
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("header", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].pageHeader },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].eyebrow }, "Admin"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h1", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].title }, spec.title),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].subtitle },
                        spec.subtitle,
                        !spec.stub && !loading ? " · " + rows.length + " rows" : ""))),
            view === "training-records" && (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].toolbar }, RECORD_SUBNAV.map(function (tab) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("button", { key: tab.id, type: "button", className: recordTab === tab.id ? _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].toolbarBtnActive : undefined, onClick: function () { return setRecordTab(tab.id); } }, tab.label)); }))),
            spec.stub ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].panel },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted }, spec.stub))) : activeListKey ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_AdminDataTable__WEBPACK_IMPORTED_MODULE_6__.AdminDataTable, { client: client, listKey: activeListKey, title: spec.title, headers: headers, columns: columns, rows: rows, loading: loading, error: error, onRefresh: function () { return setRefreshTick(function (n) { return n + 1; }); } })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].panel },
                loading && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].muted }, "Loading\u2026"),
                error && react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { className: _portal_module_scss__WEBPACK_IMPORTED_MODULE_10__["default"].error }, error))))))));
};


/***/ }),

/***/ 276:
/*!*****************************************************!*\
  !*** ./lib/shared/ui/customerPortal.module.scss.js ***!
  \*****************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
__webpack_require__(/*! ./customerPortal.module.css */ 178);
var styles = {
    root: 'root_d30fcc18',
    header: 'header_d30fcc18',
    headerLeft: 'headerLeft_d30fcc18',
    logo: 'logo_d30fcc18',
    headerCopy: 'headerCopy_d30fcc18',
    portalTitle: 'portalTitle_d30fcc18',
    welcome: 'welcome_d30fcc18',
    welcomeName: 'welcomeName_d30fcc18',
    tagline: 'tagline_d30fcc18',
    accessBadges: 'accessBadges_d30fcc18',
    badge: 'badge_d30fcc18',
    badgeOk: 'badgeOk_d30fcc18',
    badgeMuted: 'badgeMuted_d30fcc18',
    userMenu: 'userMenu_d30fcc18',
    userName: 'userName_d30fcc18',
    chevron: 'chevron_d30fcc18',
    avatar: 'avatar_d30fcc18',
    menuToggle: 'menuToggle_d30fcc18',
    navBar: 'navBar_d30fcc18',
    navItem: 'navItem_d30fcc18',
    navItemActive: 'navItemActive_d30fcc18',
    navIcon: 'navIcon_d30fcc18',
    body: 'body_d30fcc18',
    mobileWelcome: 'mobileWelcome_d30fcc18',
    stats: 'stats_d30fcc18',
    statCard: 'statCard_d30fcc18',
    statLabel: 'statLabel_d30fcc18',
    statValue: 'statValue_d30fcc18',
    statHint: 'statHint_d30fcc18',
    statHintGood: 'statHintGood_d30fcc18',
    statHintWarn: 'statHintWarn_d30fcc18',
    statIcon: 'statIcon_d30fcc18',
    statIconGreen: 'statIconGreen_d30fcc18',
    statIconOrange: 'statIconOrange_d30fcc18',
    statIconRed: 'statIconRed_d30fcc18',
    card: 'card_d30fcc18',
    cardHeader: 'cardHeader_d30fcc18',
    cardTitle: 'cardTitle_d30fcc18',
    cardActions: 'cardActions_d30fcc18',
    linkBtn: 'linkBtn_d30fcc18',
    primaryBtn: 'primaryBtn_d30fcc18',
    ghostBtn: 'ghostBtn_d30fcc18',
    searchInput: 'searchInput_d30fcc18',
    dashMain: 'dashMain_d30fcc18',
    dashSide: 'dashSide_d30fcc18',
    dashBottom: 'dashBottom_d30fcc18',
    tableWrap: 'tableWrap_d30fcc18',
    table: 'table_d30fcc18',
    candidateCell: 'candidateCell_d30fcc18',
    avatarSm: 'avatarSm_d30fcc18',
    pill: 'pill_d30fcc18',
    pillOk: 'pillOk_d30fcc18',
    pillWarn: 'pillWarn_d30fcc18',
    pillBad: 'pillBad_d30fcc18',
    tableFooter: 'tableFooter_d30fcc18',
    muted: 'muted_d30fcc18',
    error: 'error_d30fcc18',
    docGrid: 'docGrid_d30fcc18',
    docTile: 'docTile_d30fcc18',
    docIcon: 'docIcon_d30fcc18',
    docLabel: 'docLabel_d30fcc18',
    docMeta: 'docMeta_d30fcc18',
    pdfTile: 'pdfTile_d30fcc18',
    nvqList: 'nvqList_d30fcc18',
    nvqRow: 'nvqRow_d30fcc18',
    nvqTop: 'nvqTop_d30fcc18',
    nvqName: 'nvqName_d30fcc18',
    nvqCourse: 'nvqCourse_d30fcc18',
    nvqPct: 'nvqPct_d30fcc18',
    progressTrack: 'progressTrack_d30fcc18',
    progressFill: 'progressFill_d30fcc18',
    eventList: 'eventList_d30fcc18',
    eventCard: 'eventCard_d30fcc18',
    dateBadge: 'dateBadge_d30fcc18',
    dateMonth: 'dateMonth_d30fcc18',
    dateDay: 'dateDay_d30fcc18',
    eventTitle: 'eventTitle_d30fcc18',
    eventMeta: 'eventMeta_d30fcc18',
    offerGrid: 'offerGrid_d30fcc18',
    offerCard: 'offerCard_d30fcc18',
    offerBadge: 'offerBadge_d30fcc18',
    offerTitle: 'offerTitle_d30fcc18',
    offerCode: 'offerCode_d30fcc18',
    offerLink: 'offerLink_d30fcc18',
    listPage: 'listPage_d30fcc18',
    pageTitle: 'pageTitle_d30fcc18',
    pageSubtitle: 'pageSubtitle_d30fcc18',
    stubBox: 'stubBox_d30fcc18',
    bottomNav: 'bottomNav_d30fcc18',
    tabItem: 'tabItem_d30fcc18',
    tabItemActive: 'tabItemActive_d30fcc18',
    mobileNavDrawer: 'mobileNavDrawer_d30fcc18',
    hideDesktopTableCols: 'hideDesktopTableCols_d30fcc18'
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (styles);


/***/ }),

/***/ 700:
/*!******************************!*\
  !*** ./lib/shared/ui/nav.js ***!
  \******************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ADMIN_NAV: () => (/* binding */ ADMIN_NAV),
/* harmony export */   CUSTOMER_NAV: () => (/* binding */ CUSTOMER_NAV)
/* harmony export */ });
/**
 * Portal navigation — same destinations as Next.js admin / customer sidebars.
 * SPFx uses in-web-part view switching (button click) instead of App Router.
 */
var ADMIN_NAV = [
    { id: "dashboard", label: "Home" },
    { id: "companies", label: "Companies" },
    { id: "workforce", label: "Workforce" },
    { id: "training-matrix", label: "Matrix" },
    { id: "training-records", label: "Registers" },
    { id: "nvq", label: "NVQ" },
    { id: "documents", label: "Documents" },
    { id: "events", label: "Calendar" },
    { id: "offers", label: "Offers" },
    { id: "permissions", label: "Permissions" },
    { id: "automation", label: "Automation" },
    { id: "logs", label: "Audit Log" },
];
var CUSTOMER_NAV = [
    { id: "training-matrix", label: "Training Matrix" },
    { id: "dashboard", label: "Dashboard" },
    { id: "candidates", label: "Candidates / Workforce" },
    { id: "training-records", label: "Training Records" },
    { id: "nvq-progress", label: "NVQ Progress" },
    { id: "documents", label: "Documents" },
    { id: "events", label: "Events / Bookings" },
    { id: "offers", label: "Offers" },
    { id: "support", label: "Support" },
];


/***/ }),

/***/ 256:
/*!*********************************************!*\
  !*** ./lib/shared/ui/portal.module.scss.js ***!
  \*********************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
__webpack_require__(/*! ./portal.module.css */ 955);
var styles = {
    shell: 'shell_c8d82a1a',
    adminShell: 'adminShell_c8d82a1a',
    adminTopNav: 'adminTopNav_c8d82a1a',
    adminTopNavBar: 'adminTopNavBar_c8d82a1a',
    adminBrandBlock: 'adminBrandBlock_c8d82a1a',
    brand: 'brand_c8d82a1a',
    tagline: 'tagline_c8d82a1a',
    adminNavDesktop: 'adminNavDesktop_c8d82a1a',
    adminNavLink: 'adminNavLink_c8d82a1a',
    adminNavLinkActive: 'adminNavLinkActive_c8d82a1a',
    adminTopNavTrailing: 'adminTopNavTrailing_c8d82a1a',
    adminUserChip: 'adminUserChip_c8d82a1a',
    adminMenuToggle: 'adminMenuToggle_c8d82a1a',
    adminMobileScrim: 'adminMobileScrim_c8d82a1a',
    adminMobileDrawer: 'adminMobileDrawer_c8d82a1a',
    adminMobileTitle: 'adminMobileTitle_c8d82a1a',
    adminMobileNav: 'adminMobileNav_c8d82a1a',
    adminMobileLink: 'adminMobileLink_c8d82a1a',
    adminMobileLinkActive: 'adminMobileLinkActive_c8d82a1a',
    sidebar: 'sidebar_c8d82a1a',
    chip: 'chip_c8d82a1a',
    navBtn: 'navBtn_c8d82a1a',
    navBtnActive: 'navBtnActive_c8d82a1a',
    main: 'main_c8d82a1a',
    pageHeader: 'pageHeader_c8d82a1a',
    eyebrow: 'eyebrow_c8d82a1a',
    title: 'title_c8d82a1a',
    subtitle: 'subtitle_c8d82a1a',
    stats: 'stats_c8d82a1a',
    stat: 'stat_c8d82a1a',
    statValue: 'statValue_c8d82a1a',
    statLabel: 'statLabel_c8d82a1a',
    panel: 'panel_c8d82a1a',
    tableWrap: 'tableWrap_c8d82a1a',
    table: 'table_c8d82a1a',
    muted: 'muted_c8d82a1a',
    error: 'error_c8d82a1a',
    toolbar: 'toolbar_c8d82a1a',
    toolbarBtnActive: 'toolbarBtnActive_c8d82a1a',
    success: 'success_c8d82a1a',
    actionsCol: 'actionsCol_c8d82a1a',
    linkBtn: 'linkBtn_c8d82a1a',
    linkBtnDanger: 'linkBtnDanger_c8d82a1a',
    drawerBackdrop: 'drawerBackdrop_c8d82a1a',
    drawer: 'drawer_c8d82a1a',
    drawerHeader: 'drawerHeader_c8d82a1a',
    formGrid: 'formGrid_c8d82a1a',
    field: 'field_c8d82a1a',
    fieldLabel: 'fieldLabel_c8d82a1a',
    drawerFooter: 'drawerFooter_c8d82a1a',
    logoBox: 'logoBox_c8d82a1a',
    logoPreview: 'logoPreview_c8d82a1a',
    adminMobileDrawerOpen: 'adminMobileDrawerOpen_c8d82a1a',
    hubPage: 'hubPage_c8d82a1a',
    hubHero: 'hubHero_c8d82a1a',
    hubHeroInner: 'hubHeroInner_c8d82a1a',
    hubHeroEyebrow: 'hubHeroEyebrow_c8d82a1a',
    hubHeroTitle: 'hubHeroTitle_c8d82a1a',
    hubHeroSubtitle: 'hubHeroSubtitle_c8d82a1a',
    hubHeroPanel: 'hubHeroPanel_c8d82a1a',
    hubHeroPanelLabel: 'hubHeroPanelLabel_c8d82a1a',
    hubHeroPanelHint: 'hubHeroPanelHint_c8d82a1a',
    hubActionGrid: 'hubActionGrid_c8d82a1a',
    hubActionTile: 'hubActionTile_c8d82a1a',
    hubActionIcon: 'hubActionIcon_c8d82a1a',
    hubActionTitle: 'hubActionTitle_c8d82a1a',
    hubActionHint: 'hubActionHint_c8d82a1a',
    hubResources: 'hubResources_c8d82a1a',
    hubResourcesHeader: 'hubResourcesHeader_c8d82a1a',
    hubResourcesTitle: 'hubResourcesTitle_c8d82a1a',
    hubResourcesSubtitle: 'hubResourcesSubtitle_c8d82a1a',
    hubResourceGrid: 'hubResourceGrid_c8d82a1a',
    hubResourceTile: 'hubResourceTile_c8d82a1a',
    hubTone_lime: 'hubTone_lime_c8d82a1a',
    hubTone_charcoal: 'hubTone_charcoal_c8d82a1a',
    hubTone_forest: 'hubTone_forest_c8d82a1a',
    hubTone_slate: 'hubTone_slate_c8d82a1a',
    hubTone_moss: 'hubTone_moss_c8d82a1a',
    hubTone_ink: 'hubTone_ink_c8d82a1a'
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (styles);


/***/ }),

/***/ 726:
/*!***********************************************************************************************************!*\
  !*** ./node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js ***!
  \***********************************************************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   loadStyles: () => (/* binding */ loadStyles)
/* harmony export */ });
/* unused harmony exports Mode, ClearStyleOptions, configureLoadStyles, configureRunMode, flush, loadTheme, clearStyles, detokenize, splitStyles */
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
/**
 * In sync mode, styles are registered as style elements synchronously with loadStyles() call.
 * In async mode, styles are buffered and registered as batch in async timer for performance purpose.
 */
var Mode;
(function (Mode) {
    Mode[Mode["sync"] = 0] = "sync";
    Mode[Mode["async"] = 1] = "async";
})(Mode || (Mode = {}));
/**
 * Themable styles and non-themable styles are tracked separately
 * Specify ClearStyleOptions when calling clearStyles API to specify which group of registered styles should be cleared.
 */
var ClearStyleOptions;
(function (ClearStyleOptions) {
    /** only themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyThemable"] = 1] = "onlyThemable";
    /** only non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyNonThemable"] = 2] = "onlyNonThemable";
    /** both themable and non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["all"] = 3] = "all";
})(ClearStyleOptions || (ClearStyleOptions = {}));
// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
// load-themed-styles hosted on the page.
var _root = typeof window === 'undefined' ? __webpack_require__.g : window; // eslint-disable-line @typescript-eslint/no-explicit-any
// Nonce string to inject into script tag if one provided. This is used in CSP (Content Security Policy).
var _styleNonce = _root && _root.CSPSettings && _root.CSPSettings.nonce;
var _themeState = initializeThemeState();
/**
 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
 */
var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
var now = function () {
    return typeof performance !== 'undefined' && !!performance.now ? performance.now() : Date.now();
};
function measure(func) {
    var start = now();
    func();
    var end = now();
    _themeState.perf.duration += end - start;
}
/**
 * initialize global state object
 */
function initializeThemeState() {
    var state = _root.__themeState__ || {
        theme: undefined,
        lastStyleElement: undefined,
        registeredStyles: []
    };
    if (!state.runState) {
        state = __assign(__assign({}, state), { perf: {
                count: 0,
                duration: 0
            }, runState: {
                flushTimer: 0,
                mode: Mode.sync,
                buffer: []
            } });
    }
    if (!state.registeredThemableStyles) {
        state = __assign(__assign({}, state), { registeredThemableStyles: [] });
    }
    _root.__themeState__ = state;
    return state;
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load
 * event is fired.
 * @param {string | ThemableArray} styles Themable style text to register.
 * @param {boolean} loadAsync When true, always load styles in async mode, irrespective of current sync mode.
 */
function loadStyles(styles, loadAsync) {
    if (loadAsync === void 0) { loadAsync = false; }
    measure(function () {
        var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
        var _a = _themeState.runState, mode = _a.mode, buffer = _a.buffer, flushTimer = _a.flushTimer;
        if (loadAsync || mode === Mode.async) {
            buffer.push(styleParts);
            if (!flushTimer) {
                _themeState.runState.flushTimer = asyncLoadStyles();
            }
        }
        else {
            applyThemableStyles(styleParts);
        }
    });
}
/**
 * Allows for customizable loadStyles logic. e.g. for server side rendering application
 * @param {(processedStyles: string, rawStyles?: string | ThemableArray) => void}
 * a loadStyles callback that gets called when styles are loaded or reloaded
 */
function configureLoadStyles(loadStylesFn) {
    _themeState.loadStyles = loadStylesFn;
}
/**
 * Configure run mode of load-themable-styles
 * @param mode load-themable-styles run mode, async or sync
 */
function configureRunMode(mode) {
    _themeState.runState.mode = mode;
}
/**
 * external code can call flush to synchronously force processing of currently buffered styles
 */
function flush() {
    measure(function () {
        var styleArrays = _themeState.runState.buffer.slice();
        _themeState.runState.buffer = [];
        var mergedStyleArray = [].concat.apply([], styleArrays);
        if (mergedStyleArray.length > 0) {
            applyThemableStyles(mergedStyleArray);
        }
    });
}
/**
 * register async loadStyles
 */
function asyncLoadStyles() {
    // Use "self" to distinguish conflicting global typings for setTimeout() from lib.dom.d.ts vs Jest's @types/node
    // https://github.com/jestjs/jest/issues/14418
    return self.setTimeout(function () {
        _themeState.runState.flushTimer = 0;
        flush();
    }, 0);
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
 * is fired.
 * @param {string} styleText Style to register.
 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
 */
function applyThemableStyles(stylesArray, styleRecord) {
    if (_themeState.loadStyles) {
        _themeState.loadStyles(resolveThemableArray(stylesArray).styleString, stylesArray);
    }
    else {
        registerStyles(stylesArray);
    }
}
/**
 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
 * replaced.
 * @param {theme} theme JSON object of theme tokens to values.
 */
function loadTheme(theme) {
    _themeState.theme = theme;
    // reload styles.
    reloadStyles();
}
/**
 * Clear already registered style elements and style records in theme_State object
 * @param option - specify which group of registered styles should be cleared.
 * Default to be both themable and non-themable styles will be cleared
 */
function clearStyles(option) {
    if (option === void 0) { option = ClearStyleOptions.all; }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyNonThemable) {
        clearStylesInternal(_themeState.registeredStyles);
        _themeState.registeredStyles = [];
    }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyThemable) {
        clearStylesInternal(_themeState.registeredThemableStyles);
        _themeState.registeredThemableStyles = [];
    }
}
function clearStylesInternal(records) {
    records.forEach(function (styleRecord) {
        var styleElement = styleRecord && styleRecord.styleElement;
        if (styleElement && styleElement.parentElement) {
            styleElement.parentElement.removeChild(styleElement);
        }
    });
}
/**
 * Reloads styles.
 */
function reloadStyles() {
    if (_themeState.theme) {
        var themableStyles = [];
        for (var _i = 0, _a = _themeState.registeredThemableStyles; _i < _a.length; _i++) {
            var styleRecord = _a[_i];
            themableStyles.push(styleRecord.themableStyle);
        }
        if (themableStyles.length > 0) {
            clearStyles(ClearStyleOptions.onlyThemable);
            applyThemableStyles([].concat.apply([], themableStyles));
        }
    }
}
/**
 * Find theme tokens and replaces them with provided theme values.
 * @param {string} styles Tokenized styles to fix.
 */
function detokenize(styles) {
    if (styles) {
        styles = resolveThemableArray(splitStyles(styles)).styleString;
    }
    return styles;
}
/**
 * Resolves ThemingInstruction objects in an array and joins the result into a string.
 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
 */
function resolveThemableArray(splitStyleArray) {
    var theme = _themeState.theme;
    var themable = false;
    // Resolve the array of theming instructions to an array of strings.
    // Then join the array to produce the final CSS string.
    var resolvedArray = (splitStyleArray || []).map(function (currentValue) {
        var themeSlot = currentValue.theme;
        if (themeSlot) {
            themable = true;
            // A theming annotation. Resolve it.
            var themedValue = theme ? theme[themeSlot] : undefined;
            var defaultValue = currentValue.defaultValue || 'inherit';
            // Warn to console if we hit an unthemed value even when themes are provided, but only if "DEBUG" is true.
            // Allow the themedValue to be undefined to explicitly request the default value.
            if (theme &&
                !themedValue &&
                console &&
                !(themeSlot in theme) &&
                "boolean" !== 'undefined' &&
                true) {
                // eslint-disable-next-line no-console
                console.warn("Theming value not provided for \"".concat(themeSlot, "\". Falling back to \"").concat(defaultValue, "\"."));
            }
            return themedValue || defaultValue;
        }
        else {
            // A non-themable string. Preserve it.
            return currentValue.rawString;
        }
    });
    return {
        styleString: resolvedArray.join(''),
        themable: themable
    };
}
/**
 * Split tokenized CSS into an array of strings and theme specification objects
 * @param {string} styles Tokenized styles to split.
 */
function splitStyles(styles) {
    var result = [];
    if (styles) {
        var pos = 0; // Current position in styles.
        var tokenMatch = void 0;
        while ((tokenMatch = _themeTokenRegex.exec(styles))) {
            var matchIndex = tokenMatch.index;
            if (matchIndex > pos) {
                result.push({
                    rawString: styles.substring(pos, matchIndex)
                });
            }
            result.push({
                theme: tokenMatch[1],
                defaultValue: tokenMatch[2] // May be undefined
            });
            // index of the first character after the current match
            pos = _themeTokenRegex.lastIndex;
        }
        // Push the rest of the string after the last match.
        result.push({
            rawString: styles.substring(pos)
        });
    }
    return result;
}
/**
 * Registers a set of style text. If it is registered too early, we will register it when the
 * window.load event is fired.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStyles(styleArray) {
    if (typeof document === 'undefined') {
        return;
    }
    var head = document.getElementsByTagName('head')[0];
    var styleElement = document.createElement('style');
    var _a = resolveThemableArray(styleArray), styleString = _a.styleString, themable = _a.themable;
    styleElement.setAttribute('data-load-themed-styles', 'true');
    if (_styleNonce) {
        styleElement.setAttribute('nonce', _styleNonce);
    }
    styleElement.appendChild(document.createTextNode(styleString));
    _themeState.perf.count++;
    head.appendChild(styleElement);
    var ev = document.createEvent('HTMLEvents');
    ev.initEvent('styleinsert', true /* bubbleEvent */, false /* cancelable */);
    ev.args = {
        newStyle: styleElement
    };
    document.dispatchEvent(ev);
    var record = {
        styleElement: styleElement,
        themableStyle: styleArray
    };
    if (themable) {
        _themeState.registeredThemableStyles.push(record);
    }
    else {
        _themeState.registeredStyles.push(record);
    }
}


/***/ }),

/***/ 659:
/*!*****************************************!*\
  !*** ./lib/shared/assets/pave-logo.png ***!
  \*****************************************/
/***/ ((module, __unused_webpack_exports, __webpack_require__) => {

module.exports = __webpack_require__.p + "pave-logo_25884a34928b530ef7b7.png";

/***/ }),

/***/ 878:
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__878__;

/***/ }),

/***/ 272:
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__272__;

/***/ }),

/***/ 723:
/*!**********************************************!*\
  !*** external "@microsoft/sp-property-pane" ***!
  \**********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__723__;

/***/ }),

/***/ 134:
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__134__;

/***/ }),

/***/ 580:
/*!************************************************!*\
  !*** external "PaveAdminPortalWebPartStrings" ***!
  \************************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__580__;

/***/ }),

/***/ 650:
/*!************************!*\
  !*** external "react" ***!
  \************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__650__;

/***/ }),

/***/ 729:
/*!****************************!*\
  !*** external "react-dom" ***!
  \****************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__729__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	(() => {
/******/ 		var _publicPath = __RUSHSTACK_CURRENT_SCRIPT__ ? __RUSHSTACK_CURRENT_SCRIPT__.src : '';
/******/ 		__webpack_require__.p = _publicPath.slice(0, _publicPath.lastIndexOf('/') + 1);
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!****************************************************************!*\
  !*** ./lib/webparts/paveAdminPortal/PaveAdminPortalWebPart.js ***!
  \****************************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 650);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react-dom */ 729);
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react_dom__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-core-library */ 878);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-property-pane */ 723);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 134);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__);
/* harmony import */ var PaveAdminPortalWebPartStrings__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! PaveAdminPortalWebPartStrings */ 580);
/* harmony import */ var PaveAdminPortalWebPartStrings__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(PaveAdminPortalWebPartStrings__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var _shared_ui_PortalShell__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../../shared/ui/PortalShell */ 698);
/* harmony import */ var _shared_services_sharePointListService__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../../shared/services/sharePointListService */ 161);
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();








var PaveAdminPortalWebPart = /** @class */ (function (_super) {
    __extends(PaveAdminPortalWebPart, _super);
    function PaveAdminPortalWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    PaveAdminPortalWebPart.prototype.render = function () {
        var rawEmail = this.context.pageContext.user.email ||
            this.context.pageContext.user.loginName ||
            "";
        var email = (0,_shared_services_sharePointListService__WEBPACK_IMPORTED_MODULE_7__.normalizeSharePointUserEmail)(rawEmail) || rawEmail;
        var isSiteAdmin = Boolean(this.context.pageContext
            .legacyPageContext &&
            this.context.pageContext
                .legacyPageContext.isSiteAdmin);
        var element = react__WEBPACK_IMPORTED_MODULE_0__.createElement(_shared_ui_PortalShell__WEBPACK_IMPORTED_MODULE_6__.PortalShell, {
            mode: "admin",
            spHttpClient: this.context.spHttpClient,
            webUrl: this.context.pageContext.web.absoluteUrl,
            userEmail: email,
            isSiteAdmin: isSiteAdmin,
        });
        react_dom__WEBPACK_IMPORTED_MODULE_1__.render(element, this.domElement);
    };
    PaveAdminPortalWebPart.prototype.onDispose = function () {
        react_dom__WEBPACK_IMPORTED_MODULE_1__.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(PaveAdminPortalWebPart.prototype, "dataVersion", {
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__.Version.parse("1.0");
        },
        enumerable: false,
        configurable: true
    });
    PaveAdminPortalWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: PaveAdminPortalWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.PropertyPaneDescription },
                    groups: [
                        {
                            groupName: PaveAdminPortalWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.BasicGroupName,
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__.PropertyPaneTextField)("description", {
                                    label: PaveAdminPortalWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.DescriptionFieldLabel,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    };
    return PaveAdminPortalWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (PaveAdminPortalWebPart);

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});})();;
//# sourceMappingURL=pave-admin-portal-web-part.js.map