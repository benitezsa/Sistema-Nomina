'use client'

import { createContext, useCallback, useContext, useEffect, useState } from 'react'
import type { Company, UtilityParameters } from '@/components/utilities/types'

const DEFAULT_COMPANY: Company = {
  id: '1',
  name: 'EMPRESA DEMO SAC',
  ruc: '20123456789',
  activity: 'Comercio y restaurantes',
  address: 'Av. Los Héroes 123',
}

const DEFAULT_PARAMETERS: UtilityParameters = {
  income: 0,
  legalPercent: 10,
  averageEmployees: 27,
  daysBase: 360,
  remunerationMonths: 12,
  capMonths: 18,
}

type SettingsContextValue = {
  company: Company
  parameters: UtilityParameters
  setCompany: (company: Company) => void
  updateParameters: (patch: Partial<UtilityParameters>) => void
}

const SettingsContext = createContext<SettingsContextValue | null>(null)

const COMPANY_KEY = 'utilidades-company'
const PARAMETERS_KEY = 'utilidades-parameters'

function loadCompany(): Company {
  const raw = window.localStorage.getItem(COMPANY_KEY)
  if (raw) {
    const parsed = JSON.parse(raw) as Company
    if (parsed.ruc) return { ...DEFAULT_COMPANY, ...parsed } as Company
  }
  return DEFAULT_COMPANY
}

function loadParameters(): UtilityParameters {
  const raw = window.localStorage.getItem(PARAMETERS_KEY)
  if (raw) return { ...DEFAULT_PARAMETERS, ...JSON.parse(raw) } as UtilityParameters
  return DEFAULT_PARAMETERS
}

export function SettingsProvider({ children }: { children: React.ReactNode }) {
  const [company, setCompanyState] = useState<Company>(DEFAULT_COMPANY)
  const [parameters, setParametersState] = useState<UtilityParameters>(DEFAULT_PARAMETERS)

  useEffect(() => {
    setCompanyState(loadCompany())
    setParametersState(loadParameters())
  }, [])

  useEffect(() => {
    window.localStorage.setItem(COMPANY_KEY, JSON.stringify(company))
  }, [company])

  useEffect(() => {
    window.localStorage.setItem(PARAMETERS_KEY, JSON.stringify(parameters))
  }, [parameters])

  const setCompany = useCallback((next: Company) => setCompanyState(next), [])
  const updateParameters = useCallback((patch: Partial<UtilityParameters>) => {
    setParametersState((prev) => ({ ...prev, ...patch }))
  }, [])

  return (
    <SettingsContext.Provider
      value={{ company, parameters, setCompany, updateParameters }}
    >
      {children}
    </SettingsContext.Provider>
  )
}

export function useSettings(): SettingsContextValue {
  const ctx = useContext(SettingsContext)
  if (!ctx) {
    throw new Error('useSettings debe usarse dentro de un SettingsProvider')
  }
  return ctx
}