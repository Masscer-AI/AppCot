"use client";

import { useMemo, useState } from "react";
import Link from "next/link";
import { Alert, Button, Container, Paper, Stack, Text, TextInput, Title } from "@mantine/core";

export default function CotizadorVerifyPage() {
  const initialToken = useMemo(() => "", []);
  const [token, setToken] = useState(initialToken);
  const [result, setResult] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleVerify = async () => {
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8009";
      const response = await fetch(`${apiBaseUrl}/api/auth/magic-link/verify`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token: token.trim() }),
      });
      if (!response.ok) {
        throw new Error(`Error ${response.status}`);
      }
      setResult("Sesion iniciada. Ya puedes ir a /cotizaciones.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container size="sm" py={40}>
      <Stack gap="lg">
        <Title order={2}>Verificar magic link</Title>
        <Paper withBorder radius="md" p="lg">
          <Stack gap="md">
            <TextInput
              label="Token"
              placeholder="Pega el token aqui"
              value={token}
              onChange={(e) => setToken(e.currentTarget.value)}
            />
            <Button onClick={handleVerify} loading={loading}>
              Verificar e iniciar sesion
            </Button>
          </Stack>
        </Paper>

        {result ? (
          <Alert color="green" title="Exito">
            <Text size="sm">{result}</Text>
            <Text component={Link} href="/cotizaciones" size="sm">
              Ir a cotizaciones
            </Text>
          </Alert>
        ) : null}

        {error ? (
          <Alert color="red" title="Error">
            <Text size="sm">{error}</Text>
          </Alert>
        ) : null}
      </Stack>
    </Container>
  );
}
