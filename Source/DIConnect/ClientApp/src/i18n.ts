// <copyright file="i18n.ts" company="Microsoft Corporation">
// Copyright (c) Microsoft.
// Licensed under the MIT License.
// </copyright>

import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';
import Backend from 'i18next-http-backend';
import moment from 'moment';
import 'moment/min/locales.min';

export const defaultLocale = () => {
    return 'en-US';
} 

i18n
  .use(Backend)
  .use(initReactI18next)
  .init({
    fallbackLng: defaultLocale(),
    interpolation: {
      escapeValue: false, // not needed for react as it escapes by default
      }
  });

export const updateLocale = () => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const locale = params.get("locale") || defaultLocale();
    i18n.changeLanguage(locale);
    moment.locale(locale);
};

export const formatDate = (date: string) => {
    if (!date) return date;
    return moment(date).format('l LT');
}

export const formatDuration = (startDate: string, endDate: string) => {
    let result = "";
    if (startDate && endDate) {
        const difference = moment(endDate).diff(moment(startDate));
        const totalDuration = moment.duration(difference);
        // Handling the scenario of duration being more than 24 hrs as it is not done by moment.js.
        const hh = ("0" + Math.floor(totalDuration.asHours())).slice(-2);
        result = hh + moment.utc(totalDuration.asMilliseconds()).locale(defaultLocale()).format(":mm:ss")
    }
    return result;
}

export const formatNumber = (number: any) => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const locale = params.get("locale") || defaultLocale();
    return Number(number).toLocaleString(locale);
}

export default i18n;