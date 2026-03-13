"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { Alert, Button, Container, Paper, Stack, Text, TextInput, Title } from "@mantine/core";

export default function CotizadorLoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [name, setName] = useState("");
  const [token, setToken] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleRequest = async () => {
    setLoading(true);
    setError(null);
    setToken(null);
    try {
      const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8009";
      const requestResponse = await fetch(`${apiBaseUrl}/api/auth/magic-link/request`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: email.trim(), name: name.trim() || undefined }),
      });
      if (!requestResponse.ok) {
        throw new Error(`Error ${requestResponse.status}`);
      }
      const data = (await requestResponse.json()) as { magic_link_token?: string };
      const magicToken = data.magic_link_token ?? null;
      setToken(magicToken);

      if (!magicToken) {
        throw new Error("No se recibio token de autenticacion");
      }

      const verifyResponse = await fetch(`${apiBaseUrl}/api/auth/magic-link/verify`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token: magicToken }),
      });
      if (!verifyResponse.ok) {
        throw new Error(`No se pudo iniciar sesion automaticamente (${verifyResponse.status})`);
      }

      router.push("/cotizaciones");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container size="sm" py={40}>
      <Stack gap="lg">
        <Title order={2}>Ingreso de cotizadores</Title>
        <Paper withBorder radius="md" p="lg">
          <Stack gap="md">
            <TextInput label="Correo" placeholder="cotizador@empresa.com" value={email} onChange={(e) => setEmail(e.currentTarget.value)} />
            <TextInput label="Nombre (opcional)" placeholder="Tu nombre" value={name} onChange={(e) => setName(e.currentTarget.value)} />
            <Button onClick={handleRequest} loading={loading}>
              Solicitar magic link (stub)
            </Button>
          </Stack>
        </Paper>

        {token ? (
          <Alert color="green" title="Token generado">
            <Text size="sm">Inicio de sesion automatico en progreso...</Text>
            <Text size="sm" fw={700}>
              {token}
            </Text>
            <Text size="sm">Si falla, puedes usar /cotizador/verify como respaldo.</Text>
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
